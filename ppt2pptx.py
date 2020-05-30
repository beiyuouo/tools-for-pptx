import os
import argparse
import win32com
import win32com.client
from time import sleep

def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('--path', type=str, default='none', help='path to pptx')
    args = parser.parse_args()
    return args

def main():
    args = get_arguments()
    if args.path == 'none':
        exit('please enter --path for path to pptx')

    powerpoint = win32com.client.Dispatch('PowerPoint.Application')
    for subPath in os.listdir(args.path):
        subPath = os.path.join(args.path, subPath)
        if subPath.endswith('.ppt'):
            # win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
            powerpoint.Visible = 1
            ppt = powerpoint.Presentations.Open(subPath)
            print(subPath[:-4]+'.pptx')
            ppt.SaveAs(subPath[:-4]+'.pptx')
            ppt.Close()
            
            print(subPath)
            # sleep(10)
    powerpoint.Quit()

if __name__ == '__main__':
    main()