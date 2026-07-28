import sys
import os
import unittest
from unittest.mock import MagicMock, patch

# Mock 'com' and nested modules
com_mock = MagicMock()
sys.modules['com'] = com_mock
sys.modules['com.sun.star'] = com_mock.sun.star
sys.modules['com.sun.star.beans'] = com_mock.sun.star.beans

class DummyPropertyValue:
    def __init__(self, Name=None, Value=None):
        self.Name = Name
        self.Value = Value

com_mock.sun.star.beans.PropertyValue = DummyPropertyValue

# Mock 'uno' and 'unohelper' before importing LeenoUtils
uno_mock = MagicMock()
sys.modules['uno'] = uno_mock
sys.modules['unohelper'] = MagicMock()

# Mock internal modules that cause circular imports or dependencies
sys.modules['LeenoDialogs'] = MagicMock()
sys.modules['Dialogs'] = MagicMock()
sys.modules['LeenoConfig'] = MagicMock()
sys.modules['LeenoGlobals'] = MagicMock()
sys.modules['pyleeno'] = MagicMock()
sys.modules['PyPDF2'] = MagicMock()

# Append the directory of this file to sys.path
pythonpath_dir = os.path.dirname(os.path.abspath(__file__))
if pythonpath_dir not in sys.path:
    sys.path.insert(0, pythonpath_dir)

import LeenoUtils

class TestClipboardPreservation(unittest.TestCase):

    def setUp(self):
        # Reset mock for each test
        uno_mock.reset_mock()

    @patch('LeenoUtils.getComponentContext')
    def test_context_manager_saves_and_restores_clipboard(self, mock_get_context):
        # Mocking UNO components
        mock_ctx = MagicMock()
        mock_smgr = MagicMock()
        mock_clip = MagicMock()
        mock_transferable = MagicMock()

        mock_get_context.return_value = mock_ctx
        mock_ctx.getServiceManager.return_value = mock_smgr
        mock_smgr.createInstanceWithContext.return_value = mock_clip
        mock_clip.getContents.return_value = mock_transferable

        # Run context manager
        with LeenoUtils.preserve_clipboard_context():
            # Simulated operation inside the context that might overwrite clipboard
            pass

        # Verify that SystemClipboard service was requested with the context
        mock_smgr.createInstanceWithContext.assert_called_once_with(
            "com.sun.star.datatransfer.clipboard.SystemClipboard", mock_ctx
        )
        # Verify getContents was called to retrieve the original clipboard
        mock_clip.getContents.assert_called_once()
        # Verify setContents was called to restore the clipboard with original transferable
        mock_clip.setContents.assert_called_once_with(mock_transferable, None)

    @patch('LeenoUtils.getComponentContext')
    def test_decorator_preserves_clipboard(self, mock_get_context):
        mock_ctx = MagicMock()
        mock_smgr = MagicMock()
        mock_clip = MagicMock()
        mock_transferable = MagicMock()

        mock_get_context.return_value = mock_ctx
        mock_ctx.getServiceManager.return_value = mock_smgr
        mock_smgr.createInstanceWithContext.return_value = mock_clip
        mock_clip.getContents.return_value = mock_transferable

        @LeenoUtils.preserve_clipboard
        def dummy_function(x, y):
            """This is a dummy docstring."""
            return x + y

        # Verify decorator wraps metadata correctly
        self.assertEqual(dummy_function.__name__, "dummy_function")
        self.assertEqual(dummy_function.__doc__, "This is a dummy docstring.")

        # Run decorated function
        result = dummy_function(3, 4)
        self.assertEqual(result, 7)

        # Verify clipboard was saved and restored
        mock_clip.getContents.assert_called_once()
        mock_clip.setContents.assert_called_once_with(mock_transferable, None)

    @patch('LeenoUtils.getComponentContext')
    def test_restores_clipboard_on_exception(self, mock_get_context):
        mock_ctx = MagicMock()
        mock_smgr = MagicMock()
        mock_clip = MagicMock()
        mock_transferable = MagicMock()

        mock_get_context.return_value = mock_ctx
        mock_ctx.getServiceManager.return_value = mock_smgr
        mock_smgr.createInstanceWithContext.return_value = mock_clip
        mock_clip.getContents.return_value = mock_transferable

        @LeenoUtils.preserve_clipboard
        def raising_function():
            raise ValueError("Something went wrong")

        # Verify that exception propagates but clipboard is still restored
        with self.assertRaises(ValueError) as context:
            raising_function()

        self.assertEqual(str(context.exception), "Something went wrong")
        mock_clip.getContents.assert_called_once()
        mock_clip.setContents.assert_called_once_with(mock_transferable, None)

    @patch('LeenoUtils.getComponentContext')
    def test_robustness_on_uno_errors(self, mock_get_context):
        # Simulated scenario: getComponentContext raises an exception
        mock_get_context.side_effect = Exception("UNO not initialized")

        @LeenoUtils.preserve_clipboard
        def normal_function():
            return "success"

        # The function should run completely fine and not crash due to UNO clipboard issues
        result = None
        try:
            result = normal_function()
        except Exception as e:
            self.fail(f"preserve_clipboard raised an exception: {e}")

        self.assertEqual(result, "success")


if __name__ == '__main__':
    unittest.main()
