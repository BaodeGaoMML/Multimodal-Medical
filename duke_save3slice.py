import os
import pandas as pd
import pydicom
import numpy as np
from PIL import Image

import os
import pandas as pd
import pydicom
import numpy as np
from PIL import Image


from tqdm import tqdm  

 

def extract_3_slice_stack(patient_list, root_dir, anno_path, output_dir):
    # Load annotations file 
    annotations = pd.read_excel(anno_path)
    annotations.columns = annotations.columns.str.replace(' ', '')
    # Priority order for folder search based on folder names
    # first post-contrast, phase 1, subtraction of the pre-contrast from first post-contrast, 
    priority_terms = ['1st pass', 'ph1', 'post_1', 'sub', 'dyn']


    for patient_id in tqdm(patient_list, desc="Extracting Slices"):
        print(f"Processing {patient_id}...")
        
        #Get annotation coordinates based on annotation box 
        anno = annotations[annotations['PatientID'] == patient_id]
        if anno.empty:
            print(f"  X ID {patient_id} not found in Annotation file.")
            continue
        
        r1, r2, c1, c2, s1, s2 = (anno['StartRow'].iloc[0], anno['EndRow'].iloc[0], 
                                  anno['StartColumn'].iloc[0], anno['EndColumn'].iloc[0], 
                                  anno['StartSlice'].iloc[0], anno['EndSlice'].iloc[0])

        #Find the patient root folder 
        patient_base_folder = None
        normalized_target = patient_id.replace('_', '').lower()
        
   
        duke_root = os.path.join(root_dir, "Duke-Breast-Cancer-MRI")
        
        if not os.path.exists(duke_root):
            print(f"  X Path does not exist: {duke_root}")
            continue

        for folder in os.listdir(duke_root):
            if folder.replace('_', '').lower() == normalized_target:
                patient_base_folder = os.path.join(duke_root, folder)
                break
        
        if not patient_base_folder:
            print(f"  X Could not find folder for {patient_id} in {duke_root}")
            continue

        #select one study folder for each patient 
        try:
            study_folders = [f for f in os.listdir(patient_base_folder) if os.path.isdir(os.path.join(patient_base_folder, f))]
            study_path = os.path.join(patient_base_folder, study_folders[0])
        except IndexError:
            print(f"  X No study folder found inside {patient_base_folder}")
            continue

        #find the best from the study using priorities
        series_folders = os.listdir(study_path)
        selected_series_folder = None
        
        for term in priority_terms:
            for s_folder in series_folders:
                if term.lower() in s_folder.lower():
                    selected_series_folder = os.path.join(study_path, s_folder)
                    break
            if selected_series_folder: break

        if not selected_series_folder:
            print(f"  X No priority series (1st pass/ph1) found for {patient_id}")
            continue

        #Load and Sort DICOMs
        files = [f for f in os.listdir(selected_series_folder)]
        slices_data = []
        for f in files:
            file_path = os.path.join(selected_series_folder, f)
            
       
            if pydicom.misc.is_dicom(file_path):
                try:
          
                    ds = pydicom.dcmread(file_path, force=True)
                    
                 
                    if hasattr(ds, 'InstanceNumber'):
                        slices_data.append(ds)
                except Exception as e:
                    print(f"  Skipping corrupted file {f}: {e}")
        
        if not slices_data:
            print(f"  X No valid DICOM slices found in {selected_series_folder}")
            continue

        #Sort by InstanceNumber to match annotation slice numbers
        #Instance Number included in slice file metadata   
        slices_data.sort(key=lambda x: int(x.InstanceNumber))
        # Calculate 3-Slice Stack (Middle, Mid-1, Mid+1)
        mid = int(s1 + (s2 - s1) // 2)
        stack_indices = [mid - 1, mid, mid + 1]

        #save slice images
        save_path = output_dir+patient_id
        #os.makedirs(save_path, exist_ok=True)

        for s_idx in stack_indices:
            try:
                # DICOM Instance is 1-based, Python list is 0-based - changing to python indexing 
                ds = slices_data[s_idx - 1]
                pixel_array = ds.pixel_array
                
                #Crop using coordinates from Excel
                crop = pixel_array[int(r1):int(r2), int(c1):int(c2)]
                
                #Normalize and save as 8-bit png
                crop_norm = ((crop - np.min(crop)) / (np.max(crop) - np.min(crop)) * 255).astype(np.uint8)
                Image.fromarray(crop_norm).save(save_path+ f"slice_{s_idx}_stack.png")
            except Exception as e:
                print(f"  Error on stack slice {s_idx}: {e}")

        print(f"  Successfully saved stack for {patient_id}")

 


 

#Configuration
PT_LIST = "./validation_dataset_patients_list.xlsx"
ROOT ="./manifest-1654812109500"
MAP_FILE = "./Breast-Cancer-MRI-filepath_filename-mapping.xlsx"
ANNO_FILE = "./Annotation_Boxes.xlsx"
OUT = "./Extracted_Slices_Validation"


pts = pd.read_excel(PT_LIST)

patients = list(pts['PatientID'])


extract_3_slice_stack(patients, ROOT, ANNO_FILE, OUT)