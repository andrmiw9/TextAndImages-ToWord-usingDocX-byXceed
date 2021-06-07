# TextAndImages-ToWord-usingDocX-byXceed
Allows user to fill in name, place and other information, with 2 adresses of pictures in order to add this all to end of an existing Word Document

Check ReadMePls.txt for current status of a project

All this programm is based on .Net Frameword WinForms and uses DocX library made by Xceed. I tried to use the other one lib, but that library completely hadn't allowed to do anything useful for me with word file. 

Before all actions down below, check that you have your template correctly configured and all paths also (mentioned in the end of this "tutorial")

__So, how this works?__ 
# 1)First this window appears:
![First pic](https://user-images.githubusercontent.com/34866926/120944080-4a640980-c73b-11eb-8cfc-31a073136c7e.png)

Here you should choose the count of your templates you are going to take (this must have been enldess while, but project is closed, so...

# 2) Secondly, the work begins. h1 header will be used for upper header, h2 for level lower and h3 for lowest text layer.
![2](https://user-images.githubusercontent.com/34866926/120944083-4afca000-c73b-11eb-870a-de7db2e1a787.png)

# 3) Continuing👀

![3](https://user-images.githubusercontent.com/34866926/120944084-4afca000-c73b-11eb-9c0b-d0abd4e7213d.png)
You must fill in all headers (keep in mind, that h1 header will be used for storing pictures, and programm won't delete it, if u change the name), and also fill in adresses for pictures and click left (yellow) button __Refresh__. That will download pic to your PC. Button on the right called __Сохранить__ (Save) allows you to STORE the image on the disk, but do not use it in a template for adding to Word Base.

# 4) Check if it is added to Base)
![4](https://user-images.githubusercontent.com/34866926/120944085-4b953680-c73b-11eb-8b28-cac27e77c8f2.png)
That's all!

If you have troubles with paths, they are defined in file Program.cs, around line 30 and in file Form1.cs, around line 50

Also, if you are interested, there is a TODO list in file Form1.cs
