import comtypes.client
import os

def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputDir=os.path.dirname(outputFileName)
        if os.path.exists(outputDir) is False:
            os.makedirs(outputDir)
        outputFileName = os.path.splitext(outputFileName)[0] + ".pdf"
        print(outputFileName)
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()

def convert_files_in_folder(powerpoint, files):
    #files = os.listdir(folder)
    #第一个f是将files迭代出来的子元素放到 pptfiles[] 中，第二个f就是files迭代的元素
    #pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for i,file in enumerate(files):
        print(i+1, file)
        outFilePath = file.replace(mypath, mypath+"_new")
        ppt_to_pdf(powerpoint, file, outFilePath)

def foreach_folder(folder, fileList):
    files = os.listdir(folder)
    for f in files:
        ff = os.path.join(folder, f)
        if os.path.isfile(ff) and f.endswith((".ppt", ".pptx")):
            fileList.append(ff)
        elif os.path.isdir(ff):
            foreach_folder(ff, fileList)
    return fileList


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    mypath = input("请输入路径：")
    flist = foreach_folder(mypath, [])
    convert_files_in_folder(powerpoint, flist)
    powerpoint.Quit()
    
