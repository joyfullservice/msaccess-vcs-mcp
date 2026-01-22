"""Tests for VCS add-in integration."""

import os
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch
import pytest

from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration


class TestVCSAddinIntegration:
    """Tests for VCSAddinIntegration class."""
    
    def test_init_default_path(self):
        """Test initialization with default add-in path."""
        with patch.dict(os.environ, {"APPDATA": "C:\\Users\\Test\\AppData\\Roaming"}):
            addin = VCSAddinIntegration()
            expected = os.path.join(
                "C:\\Users\\Test\\AppData\\Roaming",
                "MSAccessVCS",
                "Version Control.accda"
            )
            assert addin.addin_path == expected
    
    def test_init_custom_path(self):
        """Test initialization with custom add-in path."""
        custom_path = "C:\\Custom\\Path\\addin.accda"
        addin = VCSAddinIntegration(custom_path)
        assert addin.addin_path == custom_path
    
    def test_get_default_addin_path(self):
        """Test default add-in path generation."""
        with patch.dict(os.environ, {"APPDATA": "C:\\Users\\Test\\AppData\\Roaming"}):
            addin = VCSAddinIntegration()
            path = addin._get_default_addin_path()
            assert "MSAccessVCS" in path
            assert "Version Control.accda" in path
    
    def test_verify_addin_exists_true(self, tmp_path):
        """Test verifying add-in exists."""
        # Create temporary add-in file
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        assert addin.verify_addin_exists() is True
    
    def test_verify_addin_exists_false(self):
        """Test verifying add-in does not exist."""
        addin = VCSAddinIntegration("C:\\NonExistent\\addin.accda")
        assert addin.verify_addin_exists() is False
    
    def test_load_addin_not_found(self):
        """Test loading add-in when file doesn't exist."""
        addin = VCSAddinIntegration("C:\\NonExistent\\addin.accda")
        mock_app = Mock()
        
        with pytest.raises(RuntimeError, match="VCS add-in not found"):
            addin.load_addin(mock_app)
    
    def test_load_addin_success(self, tmp_path):
        """Test successful add-in loading."""
        # Create temporary add-in file
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value=None)
        
        result = addin.load_addin(mock_app)
        
        assert result is True
        assert addin._addin_loaded is True
        assert addin._app is mock_app
        mock_app.Run.assert_called_once()
    
    def test_load_addin_com_error(self, tmp_path):
        """Test add-in loading with COM error."""
        # Create temporary add-in file
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(side_effect=Exception("COM error"))
        
        with pytest.raises(RuntimeError, match="Failed to load VCS add-in"):
            addin.load_addin(mock_app)
    
    def test_call_addin_function_not_loaded(self):
        """Test calling function when add-in not loaded."""
        addin = VCSAddinIntegration("C:\\test\\addin.accda")
        
        with pytest.raises(RuntimeError, match="VCS add-in not loaded"):
            addin._call_addin_function("TestFunction")
    
    def test_call_addin_function_success(self, tmp_path):
        """Test successful function call."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value="success")
        
        # Manually set loaded state for testing
        addin._app = mock_app
        addin._addin_loaded = True
        
        result = addin._call_addin_function("TestFunction", "arg1", "arg2")
        
        assert result == "success"
        mock_app.Run.assert_called_once()
    
    def test_call_addin_function_error(self, tmp_path):
        """Test function call with error."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(side_effect=Exception("Call failed"))
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        with pytest.raises(RuntimeError, match="Failed to call add-in function"):
            addin._call_addin_function("TestFunction")
    
    def test_get_export_folder_default(self):
        """Test default export folder path."""
        addin = VCSAddinIntegration()
        db_path = "C:\\Databases\\MyDatabase.accdb"
        
        folder = addin._get_export_folder(db_path)
        
        assert folder == "C:\\Databases\\MyDatabase.src"
    
    def test_get_export_folder_custom(self):
        """Test custom export folder path."""
        addin = VCSAddinIntegration()
        db_path = "C:\\Databases\\MyDatabase.accdb"
        custom_folder = "C:\\Custom\\Export"
        
        folder = addin._get_export_folder(db_path, custom_folder)
        
        assert folder == custom_folder
    
    def test_export_source_success(self, tmp_path):
        """Test successful source export."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value=None)
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        db_path = str(tmp_path / "test.accdb")
        result = addin.export_source(db_path)
        
        assert result["success"] is True
        assert "export_path" in result
        assert "message" in result
        mock_app.Run.assert_called()
    
    def test_export_source_failure(self, tmp_path):
        """Test export failure."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(side_effect=Exception("Export failed"))
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        db_path = str(tmp_path / "test.accdb")
        result = addin.export_source(db_path)
        
        assert result["success"] is False
        assert "Export failed" in result["message"]
    
    def test_export_vba(self, tmp_path):
        """Test VBA-only export."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value=None)
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        db_path = str(tmp_path / "test.accdb")
        result = addin.export_vba(db_path)
        
        assert result["success"] is True
        mock_app.Run.assert_called()
    
    def test_merge_build(self, tmp_path):
        """Test merge build operation."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value=None)
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        db_path = str(tmp_path / "test.accdb")
        result = addin.merge_build(db_path)
        
        assert result["success"] is True
        assert "database_path" in result
        mock_app.Run.assert_called()
    
    def test_build_from_source(self, tmp_path):
        """Test build from source operation."""
        addin_file = tmp_path / "test_addin.accda"
        addin_file.touch()
        
        addin = VCSAddinIntegration(str(addin_file))
        mock_app = Mock()
        mock_app.Run = Mock(return_value=None)
        
        addin._app = mock_app
        addin._addin_loaded = True
        
        source_folder = str(tmp_path / "source")
        result = addin.build_from_source(source_folder)
        
        assert result["success"] is True
        mock_app.Run.assert_called()
    
    def test_parse_log_file_exists(self, tmp_path):
        """Test parsing existing log file."""
        log_file = tmp_path / "Export.log"
        log_content = "Export successful\nExported 10 objects"
        log_file.write_text(log_content, encoding='utf-8')
        
        addin = VCSAddinIntegration()
        result = addin.parse_log_file(str(log_file))
        
        assert result["found"] is True
        assert result["content"] == log_content
        assert result["path"] == str(log_file)
    
    def test_parse_log_file_not_exists(self):
        """Test parsing non-existent log file."""
        addin = VCSAddinIntegration()
        result = addin.parse_log_file("C:\\NonExistent\\Export.log")
        
        assert result["found"] is False
    
    def test_parse_log_file_read_error(self, tmp_path):
        """Test parsing log file with read error."""
        # Create a file that will cause encoding error
        log_file = tmp_path / "Export.log"
        
        # Write binary content that's not valid UTF-8
        with open(log_file, 'wb') as f:
            f.write(b'\x80\x81\x82')
        
        addin = VCSAddinIntegration()
        result = addin.parse_log_file(str(log_file))
        
        # Should handle the error gracefully
        assert result["found"] is False or "error" in result


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
