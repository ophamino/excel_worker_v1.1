import os



def clean_directory(path: str):
    files = os.listdir(path)
    
    for file in files:
        os.remove(f"{path}/{file}")