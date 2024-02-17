from urllib.request import urlretrieve
import zipfile, platform
import subprocess, shutil, ctypes
import os, argparse



global bool_convert
bool_convert = {
    "y": True, 
    "n": False
    }

global system_os
global system_architecture
system_os = platform.system ()
system_architecture = platform.architecture ()

global is_admin
try:
    is_admin = os.getuid() == 0
except AttributeError:
    is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0


def get_python_executable ():

    # Windows
    if system_os == "Windows":           

        current_user = os.getlogin()
        main_python_path = f"C:\\Users\\{current_user}\\AppData\\Local\\Programs\\Python"

        python_versions = os.listdir (main_python_path)
        
        # When Python 3.10 is not installed
        if "Python310" not in python_versions:

            # Install Python 3.10
            
            if accept or bool_convert[input ("Python 3.10 is not installed. Install Now? (y/n): ")]:
                # Download Python installer
                installer_x64 = "https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
                installer_x86 = "https://www.python.org/ftp/python/3.10.11/python-3.10.11.exe"
                installer_name = "Python_installer.exe"

                if system_architecture[0] == "64bit":
                    urlretrieve(installer_x64, installer_name)
                elif system_architecture[0] == "32bit":
                    urlretrieve (installer_x86, installer_name)

                # Install Python
                os.system (".\\Python_installer.exe /passive")

                # Delete installer
                os.remove (".\\Python_installer.exe")
            
            # No to Install
            else:
                quit ()
        
        print ("Python 3.10 installed.")
        python_executable_path = f"{main_python_path}\\Python310\\python.exe"
        return python_executable_path
    
    else:
        print ("Not Supported OS")


def create_venv (python_executable_path):

    venv_name = ".\\Venv_Python310"

    if not os.path.exists (venv_name):
        print ("Building Python venv.")
        os.system (f"{python_executable_path} -m venv {venv_name}")

    if accept or bool_convert[input ("Update pip in venv? (y/n): ")]:
        os.system (f"{python_executable_path} -m pip install --upgrade pip")
    
    print ("Python venv found.")
    # Install from requirements.txt
    win_path = ".\\requirements-win-310.txt"
    linux_path = ".\\requirements-linux310.txt"

    if system_os == "Windows":
        
        # Install from Requirements
        if accept or bool_convert[input ("Install from Requirements? (y/n): ")]:
            print ("Installing Windows Requirements.")
            os.system (f"{venv_name}\\Scripts\\python.exe -m pip install -r {win_path} --no-deps")

        # Install tesserocr
        if accept or bool_convert[input ("Download and Install tesserocr? (y/n): ")]:
            tesserocr_whl = "https://github.com/simonflueckiger/tesserocr-windows_build/releases/download/tesserocr-v2.6.0-tesseract-5.3.1/tesserocr-2.6.0-cp310-cp310-win_amd64.whl"
            tesserocr_path = ".\\tesserocr-2.6.0-cp310-cp310-win_amd64.whl"
            urlretrieve(tesserocr_whl, tesserocr_path)
            os.system (f"{venv_name}\\Scripts\\python.exe -m pip install {tesserocr_path}")
            os.remove (tesserocr_path)

        # Install nostril
        if accept or bool_convert[input ("Download and Install nostril? (y/n): ")]:
            nostril_zip = "https://github.com/casics/nostril/archive/refs/heads/master.zip"
            nostril_zip_path = ".\\master_1.zip"
            nostril_unpack_path = ".\\"

            # Download nostril from Github
            if not os.path.exists (nostril_zip_path):
                urlretrieve(nostril_zip, nostril_zip_path)
            # Unzip Zip File
            with zipfile.ZipFile(nostril_zip_path,"r") as zip_ref:
                zip_ref.extractall(nostril_unpack_path)
            # Install from setup.py
            os.system (f"{venv_name}\\Scripts\\python.exe -m pip install .\\nostril-master\\.")

            # Delete Files
            os.remove (nostril_zip_path)
            shutil.rmtree(".\\nostril-master")
        
        # Install espeakng python
        if accept or bool_convert[input ("Download and Install espeakng? (y/n): ")]:
            
            # Python
            espeakng_zip = "https://github.com/sayak-brm/espeakng-python/archive/refs/heads/master.zip"
            espeakng_zip_path = ".\\master_3.zip"
            espeakng_unpack_path = ".\\"

            # Download nostril from Github
            if not os.path.exists (espeakng_zip_path):
                urlretrieve(espeakng_zip, espeakng_zip_path)
            # Unzip Zip File
            with zipfile.ZipFile(espeakng_zip_path,"r") as zip_ref:
                zip_ref.extractall(espeakng_unpack_path)
            # Install from setup.py
            os.system (f"{venv_name}\\Scripts\\python.exe -m pip install .\\espeakng-python-master\\.")

            # Delete Files
            os.remove (espeakng_zip_path)
            shutil.rmtree(".\\espeakng-python-master")

            # System
            if is_admin:
                installer_x64 = "https://github.com/espeak-ng/espeak-ng/releases/download/1.51/espeak-ng-X64.msi"
                installer_x86 = "https://github.com/espeak-ng/espeak-ng/releases/download/1.51/espeak-ng-X86.msi"
                installer_name = "espeak-ng.msi"

                if system_architecture[0] == "64bit":
                    urlretrieve(installer_x64, installer_name)
                elif system_architecture[0] == "32bit":
                    urlretrieve (installer_x86, installer_name)

                os.system (f"{installer_name} /passive")
                os.remove (f".\\{installer_name}")
                print ("Installed espeak-ng to System")

                # Add to PATH
                if system_architecture[0] == "64bit":
                    result = add_path (path="C:\Program Files\eSpeak NG")
                elif system_architecture[0] == "32bit":
                    result = add_path (path="C:\Program Files (x86)\eSpeak NG")
                
                if result.returncode == 0:
                    print("espeak-ng successfully added to [Path]")
                else:
                    print("Error while adding espeak-ng to [Path]")
            else:
                print ("Run Script as Admin and install espeak-ng again, [os install / add PATH] \n or do it yourself: https://github.com/espeak-ng/espeak-ng/releases/tag/1.51")
                    



    else: # Change Later to Linux
        print ("Installing Linux Requirements.")
        os.system (f"{venv_name}\\Scripts\\python.exe -m pip install -r {linux_path} --no-deps")
    
    
    
def install_ffmpeg ():

    # Test if ffmpeg is installed
    stdout = subprocess.run ("ffmpeg -version", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if "ffmpeg version" in str(stdout):
        ffmpeg_found =  True
    else:
        ffmpeg_found =  False
    
    # Install ffmpeg
    if not ffmpeg_found:
        ffmpeg_zip = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl-shared.zip"
        ffmpeg_zip_path = ".\\master_2.zip"
        ffmpeg_exe_path = ".\\ffmpeg-master-latest-win64-gpl-shared\\bin"
        
        # Download ffmpeg
        if not os.path.exists (".\\ffmpeg-master-latest-win64-gpl-shared"):
            if not os.path.exists (ffmpeg_zip_path):
                if accept or bool_convert[input ("ffmpeg is not downloaded. download Now? (y/n): ")]:
                    urlretrieve(ffmpeg_zip, ffmpeg_zip_path)
                
            # Unzip Zip File
            if os.path.exists (ffmpeg_zip_path):
                with zipfile.ZipFile(ffmpeg_zip_path,"r") as zip_ref:
                    zip_ref.extractall(".\\")
                
                os.remove (ffmpeg_zip_path)
            else:
                return False
        
        # Add to PATH
        result = add_path (path=ffmpeg_exe_path)
            
        if result.returncode == 0:
            print("ffmpeg successfully added to [Path]")
            return True
        else:
            print("Error while adding ffmpeg to [Path]")
            return False
    

def add_path (path):
    # manually add to path
    if not is_admin:
        print ("Skript is not Run as Admin. Run again with Admin rights or add manually to system variables:")
        print (f"Add to [Path]: {os.path.abspath(path)}")
    
    # auto add to path 
    powershell_script = f'''
    $addPath = "{os.path.abspath(path)}"
    $currentPath = [System.Environment]::GetEnvironmentVariable('PATH', [System.EnvironmentVariableTarget]::Machine)
    $newPath = "$currentPath;$addPath"
    [System.Environment]::SetEnvironmentVariable('PATH', $newPath, [System.EnvironmentVariableTarget]::Machine)
        '''
    result = subprocess.run(["powershell", "-Command", powershell_script], check=True)

    return result




def run ():
    parser = argparse.ArgumentParser(description='Install Dependencies of weeablind.')
    parser.add_argument("-y", action='store_true', help='Accept the install of everything', required=False)
    
    args = parser.parse_args()
    global accept
    accept = args.y 
    
    python_executable_path = get_python_executable ()
    create_venv (python_executable_path)
    install_ffmpeg ()



if __name__ == "__main__":
    run ()
    