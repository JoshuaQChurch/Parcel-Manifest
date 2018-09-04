from setuptools import setup
import platform

if platform.system() == "Windows":
    install_requires = [
        "pandas",
        "xlrd",
        "pyautogui"
    ]

else:
    install_requires = [
        "pandas",
        "xlrd",
        "pyobjc-core",
        "pyobjc-framework-Quartz",
        "image",
        "pyautogui"
    ]

setup(
    name = "Parcel Manifest",
    version = "1.0",
    description = "Automation tool to assist with mundane tasks.",
    author = "Joshua Church",
    author_email = "Joshua.Q.Church@gmail.com",
    url = "https://github.com/JoshuaQChurch/Parcel-Manifest",
    install_requires = install_requires
)

