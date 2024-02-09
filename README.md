# Red List Fauna Document Generator
QGis plugin that creates a .docx table of fauna with various information that are on the rote liste. Landscape planners usually do this by hand, which takes several hours. This plugin automizes the task, saving a lot of time.

## Installation
1. Download the .zip file from the main branch
2. Copy the "red_list_fauna_table" folder to your QGis plugin directory (Settings -> User profiles -> Open active profile folder -> python -> plugins). 
3. Open the Osgeo-shell and run ``pip install python-docx``
4. Restart QGis
5. (Activate plugin in plugin manager)

## Usage
1. Open Plugin

![plugin_icon](red_list_fauna_table/icon.png)

3. Choose layer

<img width="242" alt="image" src="https://github.com/Merydian/fauna-docx-qgis-plugin/assets/81414045/2d143aa4-5e6e-41b2-8225-b6a25c7cf4e4">

4. Choose fauna field

<img width="242" alt="image" src="https://github.com/Merydian/fauna-docx-qgis-plugin/assets/81414045/12dbf7c9-9279-4a16-ab3e-13ea1398fc16">

5. Choose output path

<img width="242" alt="image" src="https://github.com/Merydian/fauna-docx-qgis-plugin/assets/81414045/de49269d-e92d-4e7c-986f-48554a6a3ab0">

6. Press Ok

<img width="242" alt="image" src="https://github.com/Merydian/fauna-docx-qgis-plugin/assets/81414045/be59a48f-0dab-4121-8f59-20c12ee8989c">


## The Plugin
![Plugin UI](img/plugin.png)

## The Output
![Output](img/output.jpg)
![Legend](img/legend.jpg)



