import cx_Oracle,os,shutil,webbrowser

connectionString = r'ERIS_GIS/gis295@GMPRODC.glaciermedia.inc'
mx_aerial_path= r"\\cabcvan1fpr009\MX_aerials"
mx_aerial_metapath = r"\\cabcvan1fpr009\MX_aerials\MEXtxts"
outputPath = r"\\cabcvan1fpr004\ERIS\Additional Products\Mexico Aerials"
chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
url =r"http://gisapp.erisinfo.com/index.html?OrderId=%s"
mx_aerial_collection = os.listdir(mx_aerial_path)
mx_aerial_meta_collection = os.listdir(mx_aerial_metapath)


def getGrids():
    aerials=[]
    while True:
        try:
            grid = raw_input("Enter grid:")
            aerial = str(grid).strip().lower()
            aerials.append(aerial)
        except KeyboardInterrupt:
            break
        except ValueError:
            raise ValueError("Invalid number")
    return aerials
def pullAerial(aerial):
    tiflist = [_ for _ in mx_aerial_collection if aerial.lower() in _.lower() ]
    metalist = [_ for _ in mx_aerial_meta_collection if aerial.lower()  in _.lower()]
    if tiflist==[]:
        raise ValueError("Not Found")
    return [tiflist,metalist]

def copytoJob(tiflist,txtlist):
    jobfolder = os.path.join(outputPath,OrderNum)
    try:
        if not os.path.exists(jobfolder):
            os.makedirs(jobfolder)
        for tif in tiflist:
            shutil.copyfile(os.path.join(mx_aerial_path,tif),os.path.join(outputPath,OrderNum,tif))
        for txt in txtlist:
            shutil.copyfile(os.path.join(mx_aerial_metapath,txt),os.path.join(outputPath,OrderNum,txt))
    except OSError as e:
        raise OSError("Failed create folder")
    except IOError:
        raise IOError("Failed Copy")
    return jobfolder

def renameAerials(txtlist):
    for txt in txtlist:
        try:
            fp = open(os.path.join(folder,txt))
            for i, line in enumerate(fp):
                if i == 2:
                    for year in range(1888,2018):
                        if str(year) in line:
                            print year
                            break
                    scale = line.split(",")[0].split(":")[-1]+"000"
                    print scale
                    break
            fp.close()
            os.remove(os.path.join(folder,txt))
        except:
            raise ValueError("Metadata Failed")
        try:
            if os.path.exists(os.path.join(folder,txt[:-4]+".tif")):
                os.rename(os.path.join(folder,txt[:-4]+".tif"),os.path.join(folder,("%s_%s_%s_%s.tif")%(OrderNum,txt[:-4],scale,year)))
        except Exception as e:
            print(e)

if __name__ == '__main__':
    try:
        OrderNum = str(raw_input("Enter Order Number:")).strip()
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select order_id from orders where order_num='%s'"%(OrderNum))
        t = cur.fetchone()
        OrderID = t[0]

    finally:
        cur.close()
        con.close()


    grids = getGrids()
    print grids
    for grid in grids:
        [tifs, txts]=pullAerial(grid)
        folder = copytoJob(tifs,txts)
        renameAerials(txts)
    webbrowser.open(folder)
    webbrowser.get(chrome_path).open(url%(OrderID))

