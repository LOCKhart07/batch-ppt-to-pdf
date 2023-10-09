import comtypes.client
import os


def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint


def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType=32):
    if outputFileName[-3:] != "pdf":
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()


def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]

    failed = []

    counter = 1
    for pptfile in pptfiles:
        print(f"{counter}/{len(pptfiles)}    : {pptfile}")
        fullpath = os.path.join(folder, pptfile)
        try:
            ppt_to_pdf(powerpoint, fullpath, fullpath)
        except Exception as e:
            print(e)
            failed.append(pptfile)

        counter += 1

    print(f"List of failed files: {failed}")


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
