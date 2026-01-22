"""Security and safety validation for file operations."""

from pathlib import Path
from typing import Any


def validate_database_path(path: str) -> Path:
    """
    Validate and normalize database path.
    
    Checks:
    - File exists
    - Has valid extension (.accdb, .accda, .mdb)
    - Not a system file
    
    Args:
        path: Path to Access database file
        
    Returns:
        Resolved Path object
        
    Raises:
        ValueError: If path is invalid
    """
    db_path = Path(path).resolve()
    
    if not db_path.exists():
        raise ValueError(f"Database not found: {path}")
    
    if not db_path.is_file():
        raise ValueError(f"Path is not a file: {path}")
    
    valid_extensions = {".accdb", ".accda", ".mdb"}
    if db_path.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Invalid database extension: {db_path.suffix}. "
            f"Expected one of: {', '.join(valid_extensions)}"
        )
    
    # Don't allow system directories
    system_dirs = {"C:\\Windows", "C:\\Program Files", "C:\\Program Files (x86)"}
    for sys_dir in system_dirs:
        if str(db_path).startswith(sys_dir):
            raise ValueError(f"Cannot access database in system directory: {sys_dir}")
    
    return db_path


def validate_export_directory(path: str, allow_create: bool = True) -> Path:
    """
    Validate export directory is safe and accessible.
    
    Args:
        path: Path to export directory
        allow_create: If True, create directory if it doesn't exist
        
    Returns:
        Resolved Path object
        
    Raises:
        ValueError: If path is invalid or unsafe
    """
    export_dir = Path(path).resolve()
    
    # Don't allow system directories
    system_dirs = {"C:\\Windows", "C:\\Program Files", "C:\\Program Files (x86)"}
    for sys_dir in system_dirs:
        if str(export_dir).startswith(sys_dir):
            raise ValueError(f"Cannot export to system directory: {sys_dir}")
    
    if not export_dir.exists():
        if allow_create:
            try:
                export_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                raise ValueError(f"Cannot create export directory: {e}")
        else:
            raise ValueError(f"Export directory does not exist: {path}")
    
    if not export_dir.is_dir():
        raise ValueError(f"Path is not a directory: {path}")
    
    return export_dir


def validate_source_directory(path: str) -> Path:
    """
    Validate source directory exists and is readable.
    
    Args:
        path: Path to source directory
        
    Returns:
        Resolved Path object
        
    Raises:
        ValueError: If path is invalid
    """
    source_dir = Path(path).resolve()
    
    if not source_dir.exists():
        raise ValueError(f"Source directory does not exist: {path}")
    
    if not source_dir.is_dir():
        raise ValueError(f"Path is not a directory: {path}")
    
    return source_dir


def check_write_permission(config: dict[str, Any]) -> None:
    """
    Check if write operations are disabled.
    
    Args:
        config: Configuration dictionary with permission flags
        
    Raises:
        PermissionError: If write operations are explicitly disabled
    """
    if config.get("ACCESS_VCS_DISABLE_WRITES", False):
        raise PermissionError(
            "Write operations are disabled. "
            "Set ACCESS_VCS_DISABLE_WRITES=false (or remove it) to enable database modifications."
        )
