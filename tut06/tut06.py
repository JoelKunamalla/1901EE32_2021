import os
import re
import shutil


def regex_renamer():
	print("1. Breaking Bad")
	print("2. Game of Thrones")
	print("3. Lucifer")
	if not os.path.exists("correct_srt"):
	    os.makedirs("correct_srt")
	else:
	    pass
	webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
	season_padding = int(input("Enter the Season Number Padding: "))
	episode_padding = int(input("Enter the Episode Number Padding: "))
	if webseries_num ==1:
	    if os.path.exists("correct_srt\\Breaking Bad"):
	        shutil.rmtree("correct_srt\\Breaking Bad")
	    from distutils.dir_util import copy_tree
	    copy_tree("wrong_srt\\Breaking Bad", "correct_srt\\Breaking Bad")
	    files = os.listdir('correct_srt\\Breaking Bad')
	    for file1 in files:
	        l=[]
	        v=os.path.splitext(file1)[0]
	        vjj=os.path.splitext(file1)[1]
	        d= re.compile(r"\d+")
	        bb= re.findall(d , v)
	        l.append(bb[0][1])
	        l.append(bb[1][1])
	        b= season_padding
	        while b> 1:
	            l[0]= '0'+l[0]
	            b=b-1
	        vv= episode_padding
	        while vv> 1:
	            l[1]= '0'+l[1]
	            vv=vv-1
	        kl= f"correct_srt\\Breaking Bad\\{file1}"
	        if vjj=='.mp4':
	            kj=f"correct_srt\\Breaking Bad\\Breaking Bad season {l[0]} Episode {l[1]}{vjj}"
	        else:
	            kj=f"correct_srt\\Breaking Bad\\Breaking Bad season {l[0]} Episode {l[1]}{vjj}"
	        os.rename(kl , kj )
	elif webseries_num==2:
	    if os.path.exists("correct_srt\\Game of Thrones"):
	        shutil.rmtree("correct_srt\\Game of Thrones")
	    from distutils.dir_util import copy_tree
	    copy_tree("wrong_srt\\Game of Thrones", "correct_srt\\Game of Thrones")
	    files = os.listdir('correct_srt\\Game of Thrones')
	    for file1 in files:
	        l=[]
	        v=os.path.splitext(file1)[0]
	        vjj=os.path.splitext(file1)[1]
	        d= re.compile(r"\d+")
	        bb= re.findall(d , v)
	        l.append(bb[0][0])
	        l.append(bb[1][1])
	        b= season_padding
	        while b> 1:
	            l[0]= '0'+l[0]
	            b=b-1
	        vv= episode_padding
	        while vv> 1:
	            l[1]='0'+l[1]
	            vv=vv-1
	        dsp= v.split( '-' )
	        gol= dsp[2].split('.')
	        kl= f"correct_srt\\Game of Thrones\\{file1}"
	        if vjj=='.mp4':
	            kj=f"correct_srt\\Game of Thrones\\Game of Thrones - Season {l[0]} - Episode {l[1]} - {gol[0]}{vjj}"
	        else:
	            kj=f"correct_srt\\Game of Thrones\\Game of Thrones - Season {l[0]} - Episode {l[1]} - {gol[0]}{vjj}"
	        os.rename(kl , kj )
	elif webseries_num==3:
	    if os.path.exists("correct_srt\\Lucifer"):
	        shutil.rmtree("correct_srt\\Lucifer")
	    from distutils.dir_util import copy_tree
	    copy_tree("wrong_srt\\Lucifer", "correct_srt\\Lucifer")
	    files = os.listdir('correct_srt\\Lucifer')
	    for file1 in files:
	        l=[]
	        v=os.path.splitext(file1)[0]
	        vjj=os.path.splitext(file1)[1]
	        d= re.compile(r"\d+")
	        bb= re.findall(d , v)
	        l.append(bb[0])
	        if bb[1]=='10':
	            l.append(bb[1])
	        else:
	            l.append(bb[1][1])
	        b= season_padding
	        while b> 1:
	            l[0]='0'+l[0]
	            b=b-1
	        vv= episode_padding
	        while vv> 1:
	            l[1]='0'+l[1]
	            vv=vv-1
	        if bb[1]=='10'  and episode_padding==1:
	            l[1]='10'
	        elif bb[1]=='10':
	            l[1]=l[1][1:]
	        dsp= v.split( '-' )
	        gol= dsp[2].split('.')
	        kl= f"correct_srt\\Lucifer\\{file1}"
	        if vjj=='.mp4':
	            kj=f"correct_srt\\Lucifer\\Lucifer - Season {l[0]} - Episode {l[1]} - {gol[0]}{vjj}"
	        else:
	            kj=f"correct_srt\\Lucifer\\Lucifer - Season {l[0]} - Episode {l[1]} - {gol[0]}{vjj}"
	        os.rename(kl , kj )
	return
regex_renamer()