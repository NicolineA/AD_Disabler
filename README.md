# AD_Disabler

AD Disabler is designed to disable multiple computers by their asset or serial number.

When the script is run using run.bat, it will automatically ask for Admin rights and only run the ps1 file if you are admin.

# Pictures
## When first run you will be asked to enter a SearchBase
The script will auto fill this section in based on your current computer's distinguished name. This way you will just have to removed infromation from the begining until you are left with where you want to search for the computers you will disable.

![image](https://user-images.githubusercontent.com/56235254/126362952-79d26ce4-9cde-471a-923f-fba72eb936dd.png)

## You will then be asked for the default computer name
If all of your computers start with the same few numbers or letters this is where you would enter them. This section accepts wildcards!

![image](https://user-images.githubusercontent.com/56235254/126365780-844145a0-a8a2-4c95-802d-91c4a133231b.png)

## You are then presented with the form where you enter the computers
Computers are then entered line by line and the "Disable Computers" button is pushed. Saving the time of searching each computer individually.

After the button is pushed you are asked for each computer if you want to disable it or not.

![image](https://user-images.githubusercontent.com/56235254/126362750-22e9a545-ead8-45fb-bfbe-f884f690b028.png)
