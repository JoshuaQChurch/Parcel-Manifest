import pip
    
if __name__ == "__main__":
    requirements = ["pandas", "xlrd", "pyautogui"] 
    for req in requirements:
        pip.main(["install", req])

