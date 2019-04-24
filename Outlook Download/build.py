import os

os.system("pyinstaller --log-level=WARN \
    --onefile \
    --console \
    --icon=icon='icon\daily-stats.ico'\
    --clean\
    load.spec")