#walking a directory tree
import os
import shutil
i = 1
for foldername, subfoldername, filename in os.walk(r'C:\Users\abdoa\Desktop\Automation instgram'):
    print(f'folder name is : {foldername} ')
    print(f'subfolder name is :{foldername} \ {str(subfoldername)}')
    print(f'file name is : {foldername} \ {str(subfoldername)} \ {str(filename)}')
    print()


    for file in filename:
        if file.endswith('.mp4'):
            shutil.copy(os.path.join(foldername, file), os.path.join(foldername, str(foldername)+'-' + file ))
            i +=1
          
