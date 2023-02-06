# CAD2OpenRoadsConverter
Application created to help importing polylines from AutoCAD to OpenRoads using SnakeGrid coordinate system.

To install it, download the .zip of the project and extract it to a folder.
Then, access the folder "CAD2OpenRoadsConverter-master/dist/main" and run the main.exe file.
If you want to create a shortcut folder for the program in your desktop, go to the settings button "⚙️" and click on "Criar Atalho na área de trabalho".

To use it, create a new project or use an existing one and follow each step respectively: 

1) Use the command List in a 2D polyline on autocad and copy the coordinates to your clipboard. Then, press the button "Colar LIST 2D". 
If everything goes as intended, the program will update your .xlsx coordinates file with the filtered coordinates and set your clipboard with them in the correct formating.   

2) Use an online converter to get the SnakeGrid coordinates .csv file and then press the button "Importar SnakeGrid".
This will fill all the 2D info in your .xlsx file as well as generate an OpenRoads coordinates .txt and and the command to create a smartline from it will be copied to your clipboard.

3) On autocad, use the command List in a profile polyline of the 2D polylineyou used before to copy the coordinates to your clipboard. Then, press the button "Colar LIST Perfis".
If everything goes as intended, the program will update your .xlsx coordinates file with the filtered coordinates, create an OpenRoads profile coordinates .txt and the command to create a smartline from it will be copied to your clipboard.

4) Press the button "Obter Coordenadas 3D" to 'merge' the 2D and profile coordinates to generate a 3D coordinates. The program uses linear approximation when the points are not exactlty the same.
An OpenRoads 3D coordinates .txt file will be created and the command to create a smartline from it will be copied to your clipboard.

Last updated on: 05 feb 2023. Currently working on updates to improve data exhibition and create an .exe version of it.
