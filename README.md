# Powerpoint-Starter
Made for Centwerk Delft

## Description
The program will import all .pptx files from the current folder, and will maximise and distribute them across monitors.


## How to run
Place the built .exe in the startup folder (commonly located at : C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup) together with the .pptx files you would like to launch.
The .exe will run upon starting the user environment.

## How to build
The project is made using anaconda as the package manager.
This was due to some difficulties with the versioning of the win32api and win32gui package and the python interpreter. 
The most straightforward way to make changes to the program is thereby by using anaconda3 and creating the executable using pyinstaller.

anaconda3
https://www.anaconda.com/products/individual

pyinstaller

In the anaconda environment install the pyinstaller package and build the exe.
~~~
pip install pyinstaller
pyinstaller --noconfirm --onefile --console --exclude-module "pandas" --exclude-module "numpy"  "C:/Users/marcv/PycharmProjects/personal/automagicppt/__main__.py"
~~~
