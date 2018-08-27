import os
import pip
    
if __name__ == "__main__":
    
    if os.name == 'nt':
        requirements = ("pandas", 
                        "xlrd", 
                        "pyautogui")
    
    else:
        requirements = ("pandas", 
                        "xlrd", 
                        "pyobjc-core",
                        "pyobjc-framework-Quartz",
                        "image",
                        "pyautogui")

    for req in requirements:
        pip.main(["install", req])

