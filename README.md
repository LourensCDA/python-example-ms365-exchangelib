# Example of using exchangelib

This example tests the functionality of the exchangelib library to read MS365 exchange e=mails.

For info on the library: (https://pypi.org/project/exchangelib/)[https://pypi.org/project/exchangelib/]

To install

```
pipenv install exchangelib
```

## virtualenv

I use pipenv as virtual enviroment.

To install in Windows:

```powershell
pip install --user pipenv
```

To install in Ubuntu:

```bash
sudo apt install pipenv
```

You may have to add the Python Scripts folder to the path when running it:

Windows:

```powershell
$env:Path =  "C:\Users\~\AppData\Roaming\Python\Python39\Scripts;$env:Path"
```

You can run pipenv using the following commands in your project folder.

Windows:

```powershell
pipenv run python main.py
```

## .env

The .env file needs to contain:

```
EMLUSER=youremail@yourdomain.com
EMLPWD=yourpassword
```

Values to the right of "=" should be specific to your enviroment
