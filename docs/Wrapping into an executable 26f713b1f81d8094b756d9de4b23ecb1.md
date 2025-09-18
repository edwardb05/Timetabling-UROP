# Wrapping into an executable

Owner: Edward Brady

To complete this project I wanted it to be easily accessible for the undergrad office.  In order to do this I wanted the timetabling app to be launched from one executable. This was completed using the pyinstaller package ran on a virtual machine on github.

This is done in a GitHub action and produces an executable fro windows and Mac.

The virtual machine installs all the dependencies and then runs the command:

```bash
 run: pyinstaller --onefile --add-data "Home_Page.py;." --add-data "pages;pages" --add-data ".streamlit;.streamlit" --collect-all streamlit --collect-all pandas --collect-all ortools --collect-all rapidfuzz --collect-all openpyxl launcher.py 
```

This uses the pyinstaller module to turn [la](http://launcher.py)uncher.py into an executable and adds the relevant pages to the executable. It also makes sure the dependencies get bought into the executable.

For the executbale to run we needed to add another python file [launcher.py](http://launcher.py) this is a very simple file that just runs the code:

```python
    sys.argv = [
        "streamlit", 
        "run", app_path,
        "--server.port=8501",
        "--global.developmentMode=false"  # Disable dev mode
    ]
```

which starts the streamlit server running locally.