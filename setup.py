from setuptools import setup

install_requires = [
    "pandas",
    "xlrd",
    "pyautogui"
]

setup(
    name = "Parcel Manifest",
    version = "1.0",
    description = "Automation tool to assist with mundane tasks.",
    author = "Joshua Church",
    author_email = "Joshua.Q.Church@gmail.com",
    url = "https://github.com/JoshuaQChurch/parcel-manifest",
    install_requires = install_requires
)

