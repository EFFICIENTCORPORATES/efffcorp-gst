

from gst_utilities import reco_itr_2a

from gst_utilities import gstr2a_merge

from gst_utilities import download

import os
from tkinter import filedialog, messagebox, ttk



def main():
    print("***Welcome to GST Reconciliation tool By Efficient Corporates***")
    print("**For any error faced , please share your error in efficientcorporates.info@gmail.com**")
    print("*Comments, Queries , Feedbacks can be sent at efficientcorporates.info@gmail.com*\n")



    resp1=input("So, Do you have combined GSTR2A File? Type 'Y' or 'N'")

    if resp1.upper()=="Y":
        print("Great.!! You are ready for Reconciliation")

        resp_format=input("Have you prepared File as per the Format? Type 'Y' or 'N'")

        if resp_format.upper()=="Y":

            resp_gstr2a=input("Enter the complete file path (till extension) of the GSTR2A Combined File: ")

            resp_itr=input("Enter the complete file path (till extension) of the ITR  File: ")

            tolerance=int(input("Enter the Tolerance limit of Mismatch (say 100) : "))

            output=reco_itr_2a(resp_itr,resp_gstr2a,tolerance)

        elif resp_format.upper()=="N":

            resp4=input("We are downloading the format for you. Do you want to specify a path??Type 'Y' or 'N' ")

            if resp4.upper()=="Y":
                resp5=input("Please specify a path to a folder: ")
                output=download(pth=resp5)

            else:

                output=download()
        else:

            print("Please provide reponse as 'Y' or 'N' only ")


    elif resp1.upper()=="N":
        resp2=input("Do you want to Combine the GSTR2A? Type 'Y' or 'N' ")

        if resp2.upper()=="Y":
            resp3=input("Enter the complete file path (till extension) of any one GSTR2A file in that Folder: ")

            output=gstr2a_merge(resp3)

        else:
            print("Thanks for using the Program. All the Best")
            output="Thanks for using the Program. All the Best"

    else:
        output="Please provide reponse as 'Y' or 'N' only "


    return(output)



if __name__=='__main__':
    main()





