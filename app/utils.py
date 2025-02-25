## utility functions 
import os

##prevent overwriting if filename already ecists and instantiate a versioned name
def get_unique_filename(filepath):
    base, ext = os.path.splitext(filepath)
    counter = 1
    new_filepath = filepath
    while os.path.exists(new_filepath):
        new_filepath = f"{base} ({counter}){ext}"
        counter += 1
    return new_filepath
