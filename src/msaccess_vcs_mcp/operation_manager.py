"""
Operation manager for tracking async VBA operations.

This module provides the OperationManager class which tracks pending async
operations, routes callbacks to the correct operation queues, and provides
utilities for waiting on operation completion with progress reporting.
"""

import asyncio
import json
import logging
import os
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional

logger = logging.getLogger(__name__)

# How long an abandoned operation keeps answering "cancelled" to the add-in's poll
# when no terminal callback arrives to close it sooner.
TOMBSTONE_TTL_SECONDS = 600.0

TERMINAL_CALLBACK_TYPES = frozenset({"complete", "error", "cancelled"})


def normalize_database_path(database_path: str) -> str:
    """Key a database path so spelling, case, and relative form all match."""
    return os.path.normcase(os.path.normpath(os.path.abspath(database_path)))


class MonotonicProgressReporter:
    """Emit MCP ``notifications/progress`` with a strictly increasing sequence.

    VBA progress is scoped per category and resets (for example 28/30 queries
    then 1/50 modules). The MCP spec requires each notification's ``progress``
    value to increase, so this adapter owns the sequence and keeps VBA's
    current/total in the human-readable message instead.
    """

    def __init__(self) -> None:
        self._last: float = 0.0
        self._lock = asyncio.Lock()

    @property
    def last(self) -> float:
        return self._last

    @staticmethod
    def format_message(
        message: str = "",
        vba_progress: Any = None,
        vba_total: Any = None,
    ) -> str:
        display = message or ""
        progress = MonotonicProgressReporter._as_count(vba_progress)
        if progress is None or progress < 0:
            return display
        total = MonotonicProgressReporter._as_count(vba_total)
        if total is None or total in (0, -1):
            suffix = f"({progress:g})"
        else:
            suffix = f"({progress:g}/{total:g})"
        return f"{display} {suffix}".strip() if display else suffix

    @staticmethod
    def _as_count(value: Any) -> float | None:
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str) and value.strip():
            try:
                return float(value)
            except ValueError:
                return None
        return None

    async def emit(
        self,
        ctx: Any,
        *,
        message: str = "",
        vba_progress: Any = None,
        vba_total: Any = None,
    ) -> None:
        async with self._lock:
            self._last += 1.0
            display = self.format_message(message, vba_progress, vba_total)
            if not ctx or not hasattr(ctx, "report_progress"):
                return
            try:
                await ctx.report_progress(
                    progress=self._last,
                    total=None,
                    message=display or None,
                )
            except Exception as e:
                logger.debug(f"Could not report progress: {e}")


@dataclass
class PendingOperation:
    """Represents a pending async operation."""
    
    operation_id: str
    queue: asyncio.Queue
    database_path: Optional[str] = None  # Target database (for concurrency control)
    command: Optional[str] = None  # Command being executed (Export, Build, etc.)
    started_at: datetime = field(default_factory=datetime.now)
    timeout_ms: int = 300000  # 5 minutes default
    cancelled: bool = False  # Set to True when cancellation is requested
    
    @property
    def timeout_seconds(self) -> float:
        """Get timeout in seconds."""
        return self.timeout_ms / 1000.0
    
    @property
    def elapsed_seconds(self) -> float:
        """Get elapsed time since operation started."""
        return (datetime.now() - self.started_at).total_seconds()


class OperationManager:
    """
    Manages pending async operations and routes callbacks.
    
    This is a singleton class that tracks all pending operations across
    the MCP server process. It provides:
    - Operation registration and cleanup
    - Callback routing by operation_id
    - Async wait loop for operation completion
    
    Usage:
        manager = OperationManager.get_instance()
        
        # Register operation
        operation_id, queue = manager.register_operation()
        
        # ... start VBA operation with operation_id ...
        
        # Wait for completion
        result = await manager.wait_for_completion(operation_id, ctx)
        
        # Cleanup happens automatically on completion or error
    """
    
    _instance: Optional["OperationManager"] = None
    _lock = asyncio.Lock() if hasattr(asyncio, 'Lock') else None
    
    def __init__(self):
        """Initialize the operation manager."""
        self._operations: dict[str, PendingOperation] = {}
        self._database_operations: dict[str, str] = {}  # database_path -> operation_id
        # operation_id -> monotonic expiry. An operation nobody waits for any more is
        # still running in Access; its id stays cancelled here until Access reports
        # back or the TTL passes, so the add-in's poll can stop it.
        self._tombstones: dict[str, float] = {}
        self._loop: Optional[asyncio.AbstractEventLoop] = None
    
    @classmethod
    def get_instance(cls) -> "OperationManager":
        """Get or create the singleton instance."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    def set_event_loop(self, loop: asyncio.AbstractEventLoop) -> None:
        """Set the asyncio event loop for cross-thread queue operations."""
        self._loop = loop
    
    def register_operation(
        self,
        timeout_ms: int = 300000,
        database_path: Optional[str] = None,
        command: Optional[str] = None
    ) -> tuple[str, asyncio.Queue]:
        """
        Register a new pending operation.
        
        Args:
            timeout_ms: Timeout for the operation in milliseconds
            database_path: Target database path (for concurrency control)
            command: Command being executed (Export, Build, etc.)
            
        Returns:
            Tuple of (operation_id, queue) for receiving callbacks
        """
        operation_id = str(uuid.uuid4())
        queue: asyncio.Queue = asyncio.Queue()
        
        operation = PendingOperation(
            operation_id=operation_id,
            queue=queue,
            database_path=database_path,
            command=command,
            timeout_ms=timeout_ms
        )
        
        self._operations[operation_id] = operation
        
        # Track by database path for concurrency control
        if database_path:
            self._database_operations[normalize_database_path(database_path)] = operation_id
        
        logger.debug(f"Registered operation {operation_id} for {database_path or 'unknown'}")
        
        return operation_id, queue
    
    def unregister_operation(self, operation_id: str) -> None:
        """
        Unregister an operation (cleanup after completion).
        
        Args:
            operation_id: The operation to unregister
        """
        operation = self._operations.get(operation_id)
        if operation:
            # Clean up database tracking
            if operation.database_path:
                normalized_path = normalize_database_path(operation.database_path)
                if self._database_operations.get(normalized_path) == operation_id:
                    del self._database_operations[normalized_path]
            
            del self._operations[operation_id]
            logger.debug(f"Unregistered operation {operation_id}")

    def abandon_operation(self, operation_id: str) -> None:
        """
        Stop tracking an operation nobody is waiting for, and ask Access to stop it.

        Called when a wait times out or the client gives up. The operation is still
        running in Access, so its id is kept as a tombstone that keeps answering
        "cancelled" to the add-in's poll.
        """
        if operation_id not in self._operations:
            return
        self._operations[operation_id].cancelled = True
        self._tombstones[operation_id] = time.monotonic() + TOMBSTONE_TTL_SECONDS
        self.unregister_operation(operation_id)
        logger.info(f"Abandoned operation {operation_id}; cancellation requested")

    def _tombstone_is_live(self, operation_id: str) -> bool:
        expiry = self._tombstones.get(operation_id)
        if expiry is None:
            return False
        if time.monotonic() >= expiry:
            self._tombstones.pop(operation_id, None)
            return False
        return True

    def list_active_operations(self) -> list[dict[str, Any]]:
        """Describe every operation a tool is still waiting on."""
        return [
            {
                "operation_id": op.operation_id,
                "command": op.command,
                "database_path": op.database_path,
                "started_at": op.started_at.isoformat(),
                "elapsed_seconds": round(op.elapsed_seconds, 1),
                "cancelled": op.cancelled,
            }
            for op in list(self._operations.values())
        ]
    
    def route_callback(self, operation_id: str, data: dict) -> bool:
        """
        Route a callback to the appropriate operation queue.
        
        This is called from the HTTP callback handler thread.
        Uses call_soon_threadsafe to safely put items on the async queue.
        
        Args:
            operation_id: The operation ID from the callback
            data: The callback data (type, progress, message, etc.)
            
        Returns:
            True if callback was routed, False if operation not found
        """
        operation = self._operations.get(operation_id)
        if not operation:
            if operation_id in self._tombstones:
                # The abandoned operation has finished; stop answering "cancelled".
                if data.get("type") in TERMINAL_CALLBACK_TYPES:
                    self._tombstones.pop(operation_id, None)
                return False
            logger.warning(f"Callback for unknown operation: {operation_id}")
            return False
        
        # Put callback on queue - need thread-safe approach
        if self._loop:
            # Called from HTTP thread - use threadsafe method
            self._loop.call_soon_threadsafe(
                operation.queue.put_nowait,
                data
            )
        else:
            # Called from async context - put directly
            try:
                operation.queue.put_nowait(data)
            except Exception as e:
                logger.error(f"Failed to queue callback: {e}")
                return False
        
        logger.debug(f"Routed callback to operation {operation_id}: {data.get('type')}")
        return True
    
    def get_operation(self, operation_id: str) -> Optional[PendingOperation]:
        """Get a pending operation by ID."""
        return self._operations.get(operation_id)
    
    def get_active_operation_for_database(self, database_path: str) -> Optional[PendingOperation]:
        """
        Get the active operation for a database, if any.
        
        Args:
            database_path: Path to the database
            
        Returns:
            The active PendingOperation, or None if no operation is running
        """
        operation_id = self._database_operations.get(normalize_database_path(database_path))
        if operation_id:
            return self._operations.get(operation_id)
        return None
    
    def is_database_busy(self, database_path: str) -> bool:
        """
        Check if a database has an operation in progress.
        
        Args:
            database_path: Path to the database
            
        Returns:
            True if an operation is currently running for this database
        """
        return self.get_active_operation_for_database(database_path) is not None
    
    def get_busy_status(self, database_path: str) -> Optional[dict[str, Any]]:
        """
        Get busy status information for a database.
        
        Args:
            database_path: Path to the database
            
        Returns:
            Dict with busy status info, or None if not busy
        """
        operation = self.get_active_operation_for_database(database_path)
        if operation:
            return {
                "busy": True,
                "operation_id": operation.operation_id,
                "command": operation.command,
                "started_at": operation.started_at.isoformat(),
                "elapsed_seconds": operation.elapsed_seconds,
                "message": f"Database is busy: {operation.command or 'operation'} in progress"
            }
        return None
    
    def is_cancelled(self, operation_id: str) -> bool:
        """
        Check if an operation has been cancelled.
        
        This is called from the HTTP server to respond to VBA polling.
        
        Args:
            operation_id: The operation to check
            
        Returns:
            True if cancellation was requested
        """
        operation = self._operations.get(operation_id)
        if operation:
            return operation.cancelled
        # An abandoned operation stays cancelled; unknown ones never were.
        return self._tombstone_is_live(operation_id)
    
    def request_cancel(self, operation_id: str) -> bool:
        """
        Request cancellation of an operation.
        
        Sets the cancelled flag on the operation. VBA will detect this
        when it polls the /cancel-status endpoint.
        
        Args:
            operation_id: The operation to cancel
            
        Returns:
            True if operation was found and marked for cancellation
        """
        operation = self._operations.get(operation_id)
        if operation:
            operation.cancelled = True
            logger.info(f"Cancellation requested for operation {operation_id}")
            return True
        if self._tombstone_is_live(operation_id):
            return True
        
        logger.warning(f"Cannot cancel unknown operation: {operation_id}")
        return False
    
    def create_callback_info(
        self,
        operation_id: str,
        callback_url: str,
        client: str = "cursor"
    ) -> str:
        """
        Create the callback info JSON string for VBA.
        
        Args:
            operation_id: The operation ID
            callback_url: The callback server URL
            client: Client identifier
            
        Returns:
            JSON string for VBA APIAsync first parameter
        """
        info = {
            "callback_url": callback_url,
            "operation_id": operation_id,
            "client": client,
            # Tells the add-in this server answers /cancel-status from its own
            # thread, so polling it cannot stall the operation.
            "cancel_poll": True,
        }
        return json.dumps(info)
    
    async def wait_for_completion(
        self,
        operation_id: str,
        ctx: Any = None,
        timeout_seconds: Optional[float] = None
    ) -> dict[str, Any]:
        """
        Wait for an async operation to complete.
        
        Processes callbacks from the queue, forwarding progress to the MCP
        context if provided. Returns when a 'complete' or 'error' callback
        is received, or when timeout occurs.
        
        Args:
            operation_id: The operation to wait for
            ctx: Optional MCP context for progress reporting
            timeout_seconds: Override timeout (uses operation's timeout if None)
            
        Returns:
            Dict with success status and any result/error message
        """
        operation = self._operations.get(operation_id)
        if not operation:
            return {"success": False, "error": f"Unknown operation: {operation_id}"}
        
        timeout = timeout_seconds or operation.timeout_seconds
        
        # Collect log messages to include in result
        log_messages: list[str] = []
        reporter = MonotonicProgressReporter()
        
        try:
            async with asyncio.timeout(timeout):
                while True:
                    callback = await operation.queue.get()
                    
                    msg_type = callback.get("type", "")
                    progress = callback.get("progress")
                    total = callback.get("total")
                    message = callback.get("message", "")
                    
                    if msg_type == "progress":
                        await reporter.emit(
                            ctx,
                            message=message,
                            vba_progress=progress,
                            vba_total=total,
                        )
                        logger.debug(f"Progress: {progress}/{total} - {message}")
                        
                    elif msg_type == "log":
                        # Log message from VBA - collect for the result and
                        # surface via monotonic progress (MCP logging notifications
                        # are deprecated).
                        if message:
                            log_messages.append(message)
                            await reporter.emit(ctx, message=message)
                        logger.info(f"VBA log: {message}")
                        
                    elif msg_type == "complete":
                        # Operation completed successfully
                        logger.info(f"Operation {operation_id} completed: {message}")
                        return {
                            "success": True,
                            "message": message,
                            "result": callback.get("result"),
                            "log_path": callback.get("log_path"),
                            "results_path": callback.get("results_path"),
                            "log_messages": log_messages if log_messages else None
                        }
                        
                    elif msg_type == "error":
                        # Operation failed. Test runs still attach results_path
                        # here: failed assertions complete as eorFailed.
                        logger.error(f"Operation {operation_id} failed: {message}")
                        return {
                            "success": False,
                            "error": message,
                            "code": callback.get("code"),
                            "log_path": callback.get("log_path"),
                            "results_path": callback.get("results_path"),
                            "result": callback.get("result"),
                        }
                    
                    elif msg_type == "cancelled":
                        # Operation was cancelled
                        logger.info(f"Operation {operation_id} cancelled: {message}")
                        return {
                            "success": False,
                            "cancelled": True,
                            "error_pattern": "operation_cancelled",
                            "message": message or "Operation cancelled",
                            "error": message or "Operation cancelled",
                            "log_path": callback.get("log_path"),
                            "results_path": callback.get("results_path"),
                            "result": callback.get("result"),
                        }
                    
                    else:
                        logger.warning(f"Unknown callback type: {msg_type}")
                        
        except asyncio.TimeoutError:
            logger.error(f"Operation {operation_id} timed out after {timeout}s")
            self.abandon_operation(operation_id)
            return {
                "success": False,
                "error": (
                    f"Operation timed out after {timeout} seconds. "
                    "The add-in was asked to stop it."
                ),
                "error_pattern": "timeout",
                "operation_id": operation_id,
            }
        except asyncio.CancelledError:
            # The client stopped waiting; Access would otherwise run on unobserved.
            self.abandon_operation(operation_id)
            raise
        finally:
            # A no-op after abandon_operation, which already unregistered it.
            self.unregister_operation(operation_id)
    
    def pending_count(self) -> int:
        """Get the number of pending operations."""
        return len(self._operations)
