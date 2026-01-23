"""
HTTP callback server for receiving progress updates from VBA.

This module provides a lightweight HTTP server that runs in a background thread,
receiving callbacks from the Access VBA add-in during long-running operations.
Callbacks are routed to the appropriate async operation queue based on operation_id.
"""

import json
import logging
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)


class CallbackHandler(BaseHTTPRequestHandler):
    """
    HTTP request handler for VBA callbacks.
    
    Handles:
    - POST /callback - progress updates, completion, errors, cancellation
    - GET /cancel-status/{operation_id} - check if operation should be cancelled
    - POST /cancel/{operation_id} - request cancellation (from VBA or MCP)
    - GET /health - health check
    """
    
    # Reference to callback router function (set by CallbackServer)
    callback_router: Optional[Callable[[str, dict], None]] = None
    # Reference to cancel status checker (set by CallbackServer)
    cancel_checker: Optional[Callable[[str], bool]] = None
    # Reference to cancel requester (set by CallbackServer)
    cancel_requester: Optional[Callable[[str], bool]] = None
    
    def log_message(self, format: str, *args) -> None:
        """Suppress default HTTP logging - use our logger instead."""
        logger.debug(f"HTTP: {format % args}")
    
    def do_POST(self) -> None:
        """Handle POST requests."""
        # POST /callback - receive progress/completion callbacks
        if self.path == "/callback":
            self._handle_callback()
            return
        
        # POST /cancel/{operation_id} - request cancellation
        if self.path.startswith("/cancel/"):
            operation_id = self.path[8:]  # Remove "/cancel/" prefix
            if operation_id:
                self._handle_cancel_request(operation_id)
            else:
                self._send_response(400, {"error": "Missing operation_id"})
            return
        
        self._send_response(404, {"error": "Not found"})
    
    def _handle_callback(self) -> None:
        """Handle POST /callback for progress updates."""
        try:
            # Read and parse JSON body
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length == 0:
                self._send_response(400, {"error": "Empty request body"})
                return
            
            body = self.rfile.read(content_length)
            data = json.loads(body.decode("utf-8"))
            
            # Validate required fields
            operation_id = data.get("operation_id")
            if not operation_id:
                self._send_response(400, {"error": "Missing operation_id"})
                return
            
            msg_type = data.get("type")
            if not msg_type:
                self._send_response(400, {"error": "Missing type"})
                return
            
            # Route callback to appropriate operation queue
            if self.callback_router:
                self.callback_router(operation_id, data)
                self._send_response(200, {"received": True})
            else:
                logger.warning("No callback router configured")
                self._send_response(500, {"error": "Server not configured"})
                
        except json.JSONDecodeError as e:
            logger.warning(f"Invalid JSON in callback: {e}")
            self._send_response(400, {"error": f"Invalid JSON: {e}"})
        except Exception as e:
            logger.error(f"Error processing callback: {e}")
            self._send_response(500, {"error": str(e)})
    
    def _handle_cancel_request(self, operation_id: str) -> None:
        """Handle POST /cancel/{operation_id} to request cancellation."""
        try:
            if self.cancel_requester:
                success = self.cancel_requester(operation_id)
                self._send_response(200, {
                    "operation_id": operation_id,
                    "cancelled": success,
                    "message": "Cancellation requested" if success else "Operation not found"
                })
            else:
                logger.warning("No cancel requester configured")
                self._send_response(500, {"error": "Server not configured"})
        except Exception as e:
            logger.error(f"Error processing cancel request: {e}")
            self._send_response(500, {"error": str(e)})
    
    def do_GET(self) -> None:
        """Handle GET requests."""
        # GET /health - health check
        if self.path == "/health":
            self._send_response(200, {"status": "ok"})
            return
        
        # GET /cancel-status/{operation_id} - check if cancelled
        if self.path.startswith("/cancel-status/"):
            operation_id = self.path[15:]  # Remove "/cancel-status/" prefix
            if operation_id:
                self._handle_cancel_status(operation_id)
            else:
                self._send_response(400, {"error": "Missing operation_id"})
            return
        
        self._send_response(404, {"error": "Not found"})
    
    def _handle_cancel_status(self, operation_id: str) -> None:
        """Handle GET /cancel-status/{operation_id} to check cancellation."""
        try:
            if self.cancel_checker:
                is_cancelled = self.cancel_checker(operation_id)
                self._send_response(200, {
                    "operation_id": operation_id,
                    "cancelled": is_cancelled
                })
            else:
                # No checker configured - assume not cancelled
                self._send_response(200, {
                    "operation_id": operation_id,
                    "cancelled": False
                })
        except Exception as e:
            logger.error(f"Error checking cancel status: {e}")
            self._send_response(500, {"error": str(e)})
    
    def _send_response(self, status_code: int, data: dict) -> None:
        """Send JSON response."""
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode("utf-8"))


class CallbackServer:
    """
    HTTP server for receiving VBA callbacks.
    
    Runs in a background thread, listening on a dynamically allocated port.
    Each MCP server process gets its own CallbackServer instance with a unique port.
    
    Usage:
        server = CallbackServer(callback_router=my_router_func)
        server.start()
        print(f"Callback URL: {server.callback_url}")
        # ... later ...
        server.stop()
    """
    
    _instance: Optional["CallbackServer"] = None
    _lock = threading.Lock()
    
    def __init__(
        self,
        callback_router: Callable[[str, dict], None],
        cancel_checker: Optional[Callable[[str], bool]] = None,
        cancel_requester: Optional[Callable[[str], bool]] = None,
        host: str = "127.0.0.1",
        port: int = 0  # 0 = OS assigns available port
    ):
        """
        Initialize the callback server.
        
        Args:
            callback_router: Function to route callbacks by operation_id.
                             Signature: (operation_id: str, data: dict) -> None
            cancel_checker: Function to check if operation is cancelled.
                            Signature: (operation_id: str) -> bool
            cancel_requester: Function to request operation cancellation.
                              Signature: (operation_id: str) -> bool
            host: Host to bind to (default: localhost only for security)
            port: Port to bind to (default: 0 for OS-assigned)
        """
        self.host = host
        self.port = port
        self.callback_router = callback_router
        self.cancel_checker = cancel_checker
        self.cancel_requester = cancel_requester
        self._server: Optional[HTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._running = False
    
    @classmethod
    def get_instance(cls) -> Optional["CallbackServer"]:
        """Get the singleton instance if it exists."""
        return cls._instance
    
    def start(self) -> None:
        """Start the HTTP server in a background thread."""
        if self._running:
            logger.warning("Callback server already running")
            return
        
        # Create handler class with router and cancel function references
        handler_class = type(
            "ConfiguredCallbackHandler",
            (CallbackHandler,),
            {
                "callback_router": self.callback_router,
                "cancel_checker": self.cancel_checker,
                "cancel_requester": self.cancel_requester,
            }
        )
        
        # Create and bind server (port 0 = OS assigns available port)
        self._server = HTTPServer((self.host, self.port), handler_class)
        
        # Get the actual port if we requested 0
        actual_port = self._server.server_address[1]
        self.port = actual_port
        
        # Start server thread
        self._thread = threading.Thread(
            target=self._serve_forever,
            name="CallbackServer",
            daemon=True  # Thread dies when main process exits
        )
        self._running = True
        self._thread.start()
        
        # Set singleton instance
        with self._lock:
            CallbackServer._instance = self
        
        logger.info(f"Callback server started on {self.host}:{self.port}")
    
    def _serve_forever(self) -> None:
        """Run the server loop (called in background thread)."""
        try:
            self._server.serve_forever()
        except Exception as e:
            logger.error(f"Callback server error: {e}")
        finally:
            self._running = False
    
    def stop(self) -> None:
        """Stop the HTTP server."""
        if not self._running:
            return
        
        self._running = False
        
        if self._server:
            self._server.shutdown()
            self._server.server_close()
            self._server = None
        
        if self._thread:
            self._thread.join(timeout=5.0)
            self._thread = None
        
        with self._lock:
            if CallbackServer._instance is self:
                CallbackServer._instance = None
        
        logger.info("Callback server stopped")
    
    @property
    def callback_url(self) -> str:
        """Get the full callback URL for this server."""
        return f"http://{self.host}:{self.port}/callback"
    
    @property
    def is_running(self) -> bool:
        """Check if the server is running."""
        return self._running
