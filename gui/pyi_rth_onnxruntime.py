import sys
import os
import traceback
import ctypes # Needed for manual load

# Minimal hook focusing on DLL loading

if sys.platform == 'win32':
    try:
        # Check for _MEIPASS first
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
            onnx_capi_dir = os.path.join(base_path, 'onnxruntime', 'capi')

            # Add relevant directories to DLL search path
            # It seems both might be needed in some complex cases
            os.add_dll_directory(base_path)
            os.add_dll_directory(onnx_capi_dir)
            
            # Attempt manual load using ctypes - this seems necessary to resolve initialization issues
            onnx_dll_path = os.path.join(onnx_capi_dir, 'onnxruntime.dll')
            onnx_shared_dll_path = os.path.join(onnx_capi_dir, 'onnxruntime_providers_shared.dll')
            
            # Load shared providers first if it exists (might be optional depending on ORT version/build)
            if os.path.exists(onnx_shared_dll_path):
                 ctypes.CDLL(onnx_shared_dll_path)
                 
            # Load main ORT DLL
            if os.path.exists(onnx_dll_path):
                 ctypes.CDLL(onnx_dll_path)
            else:
                 # If main DLL isn't found, this hook likely can't fix it anyway
                 pass 
                 
    except Exception:
        # In case of error during hook execution, just print to original stderr if possible
        # or ignore, as logging might not be set up.
        traceback.print_exc() 
        pass # Avoid crashing the hook itself 