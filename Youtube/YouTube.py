from pytube import Playlist, YouTube
import os

#for print
#------------------
import sys
import time
def delay_print(s):
    for c in s:
        sys.stdout.write(c)
        sys.stdout.flush()
        time.sleep(0.01)
#--------------
#check sheet (1)

delay_print("""
Automation BY:  
                    ---------------      Abdelrahman Ataa       ---------------
______________________________________________________________________________________________________________


""")
















folder_path = input("Enter your Folder Path: \n")
print("\n******************************************************************************************")
url_path = input("Enter playlist url: \n")




playlist_url = url_path
p = Playlist(playlist_url)
downloaded = []
direction = folder_path
os.chdir(direction)
for url in p.video_urls:
    try:
        yt = YouTube(url)
    except Exception:
        print(f'Video {url} is unavaialable, skipping.')
    else:
        if url in downloaded:
                print(f'Downloaded: {url}')
        else:
            gg = direction+"\\"
            print(f'Downloading: {url}')   
            yt.streams.get_highest_resolution().download(gg)
            downloaded.append(url)
            print(downloaded)


