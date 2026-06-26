import os
import sys

if sys.platform == 'win32':
    cuda_base = r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA"
    if os.path.exists(cuda_base):
        for folder in os.listdir(cuda_base):
            for sub in ["bin", r"bin\x64"]:
                full_path = os.path.join(cuda_base, folder, sub)
                if os.path.isdir(full_path):
                    try:
                        os.add_dll_directory(full_path)
                    except Exception:
                        pass

import cv2
print(cv2.__version__)
print(cv2.cuda.getCudaEnabledDeviceCount())
