# Efficient Corporates GST Utilities in Python
First of all, Welcome to Efficient Corporates..! 
Before you start reading the Documentation, let us please give you a small Disclaimer


*At, Efficient Corporates (or Eff Corp as we call it), we believe that Coding is for everyone and 
everyone has the right to make their life simpler through use of automation.
Hence, we endeavor to make it simpler for non coders (like the people without Computer Science as their Major
even Commerce / Arts  background people)to make coding a daily habit . Since,
we believe on making things less technical and more practial, hence, it is possible that 
we may not follow the best coding practices or documentation practice to ensure that 
users get to understand easily , So.. PLEASE EXCUSE US  for that*



####Lets know a bit about Efficient Corporates

Efficient Corporates is an Open Knowledge Sharing Platform which works towards 
encouraging the HR, Sales , Opertions , Finance , Tax , Accounts or any other Professionals to 
adopt technology in their day to day office Working and automate their most boring repetitive
tasks at Office (so that they can have a meaningful 9 to 5 job or WFH)

We have come up with our First Module on GST Automation and we plan to come up with many more such 
modules.

*But Hey.. We need your help on this....If you have any task , you hate doing, and wish that could be 
automated, feel free to send a mail at efficientcorporates.info@gmail.com and we will love to connect)*


###Okay, Back to this GST Module...

This Module of GST (is just like any any other python module like Pandas , Numpy, matplotlib) and has
some functions (pre defined actions) which help us to automate the GST related tasks.

Just to understand in simple terms, Let's assume you need to do the GSTR2A reconciliation, 
now the first thing you would need is a combined GSTR2A .

So, we all know , how to combine monthly GSTR2A in Excel, but, if the same needs to be done in python,
it needs a long coding for it. So, now, in this module , what we have done is , we converted that 
long coding in a small function named "gstr2a_merge" which takes just one parameter or just one 
input i.e filepath (full path to any one file where the monthly GSTR2A are stored) .

Just giving this one line of code , will make the entire block of code run in the background and 
you will get the desired output in that folder.

So, similar to this gstr2a_merge, there are many other Utilities present in this GST Module by Eff Corp:


*Again, if you are very new to Python, and have no Idea what's going on here, we strongly recommend you to 
first go through our very basic Python Tuitorial Especially designed for Finance Professionals and students*

[Python Tutorial for Absolute beginners by Effcient Corporates](https://www.youtube.com/watch?v=E509BVUxrZg&list=PLaso8-OZjhbyTgqcLSxbusK2RpPu_c3lC):



###Utilities Present in the Current Latest Version of Eff Corp GST

1. Monthly GSTR2A Merging into Single Combined File   (*Function name ***gstr2a_merge****)
2. Reconciliation of GSTR2A with the Purchase Register  (*Function name ***reco_itr_2a****)
3. GST Number Check Sum Verfication   (*Function name ***gstchecksum****)
4. Extracting PAN number from given GST input   (*Function name ***extract_pan****)
5. Checking whether invoice number is Valid or not   (*Function name ***gstinvcheck****)



Yes, there are a few very basic functions as well, like Extracting PAN , which can easily be done 
through indexing, but, we hve intentionally included these, so that the users can understand
how functions can be written even in 2-3 lines of code and not get overwhelmed by looking only at functions like
reco_itr_2a & gstr2a_merge which are over a thousand lines of python coding.


## Installation

As we mentioned that effcorp-gst is like any other Python Modules like pandas, matplotlib, numpy,
so, even this can be installed using the simple pip command as below. (Documentation available at [PyPI](https://pypi.org/project/effcorp-gst/):)

```pip install effcorp-gst```

The Efficient Corporates GST Tool will better run on Python 3.6 and above.


## How to use?

Now, lets see how do we use Pandas . We simply pip install it and then import the module saying


```import pandas as pd```

But, there are some modules like Tkinter, where we write as 

```from tkinter import *```

So, in this case , though we have pip install effcorp-gst , but the entry point to this module is
set to the keyword "gst".

This means you can enter inside this module using the Keyword "gst" only.

Quite weird, but that is the way the module has been set up, So we will need to do the below to get inside the module

```import gst```

#####OR

```from gst import gst_utilities```


***Below codes to import will not work***

```import effcorp-gst``` >>> Won't work


```from effcorp-gst import gst``` >>> Won't work



###Utilities Under this GST Tool

The structure of the module is like below.

We have a gst folder, and inside that there is a gst_utilities.py file. Inside this .py file, we have the 
various functions, like gstr2a_merge , or reco_itr_2a. So, we need to access these functions inside
utilities files.

Lets talk about each utilities inside this module.



###1. Monthly GSTR2A Merging into Single Combined File

After installing the Module through Pip, you can import the module

The name for calling the module is gst, so use below to call the module

```
from gst import gst_utilities

gst_utilities.gstr2a_merge(complete_filepath_to_gstr2a_file)
```


Just executing this code will provide you the desired Combined GSTR2A excel file in that same Folder

You will have to note the below aspects:

a. All the Monthly GSTR2A should be downloaded from GST Portal site and not from any other site or package software.
(This is because the format of the GSTR2A is very critical for this function) Format should be same 
as is available in the GST Site .

b. All the files must be a .xlsx format and not a zip file or other format

c. All excel files you wish to combine , must be inside a single folder.

d. The input parameter is complete path to any one excel file in that folder. for examples
you have 12 excel files inside a folder GSTR2A which is in desktop. So, you will provide the complete path to 
any one file which is insdie this GSTR2A folder. Something like this.

***C:\Desktop\GSTR2A\April.xlsx***

So, giving like this will read all other excel files automatically and store the output file in this folder GSTR2A.



###2. Reconciliation of GSTR2A and the ITR

So, GSTR2A reco has always been a major issue for most of the practicing professionals.
Here we present a function of python, which will compare the GSTR2A and the Purchase Register
and will give you the matched and unmatched data.


```
from gst import gst_utilities

gst_utilities.reco_itr_2a(path_to_itr, path_to_consolidated_gstr2a , tolerance limit)
```
This function takes the 3 parameters.First Two are Mandatory and 1 is optional

path to ITR : This argument should be the complete path to the ITR file which is as per the format .
            Please ensure to provide the complete filepath of ITR till the extension

path_to_consolidated_gstr2a : This is the argument for the complete filepath of the GSTR2A file.             
                            Please ensure to gve the complete file path till the extension

tolerance limit : This is also next important parameter. This is the Tolerance limit.
                    If a invoice is booked with Tax of Rs 12,300 , but the same invoice is given in GSTR2A as Rs 12450.
                    Now, there is a difference of Rs 150. Now , if the tolerance limit is kept as 100, then this case will be considered NOT MATCHING
                    But, if the tolerance limit is kept as 200, then this case will be considered as a match
                    Use can provide the Tolerance limit value based on the size of the client
                    If no parameter is provided , then the 100 is taken as the Tolerance limit


###3. Check Sum Validation for GST Number

We know that the 15 digit GST Number is consists of 

First 2 digit : State Code

Next 10 digit : PAN Number

Next 1 digit : No of Entity in Same PAN in that State

Next 1 digit : "Z"

Last 1 digit : Check Sum

This check sum helps us to identify whether the GST Number is Valid or not.

So, Provide the Input parameter as GST Number and it will return one of the below:

a. Check Sum MATCH
b. Check Sum MISMATCH

To use the function

```
from gst import gst_utilities

gst_utilities.gstchecksum("07AAAAT7798M2ZK")
```

Output:
'Check Sum MATCH'

```
from gst import gst_utilities

gst_utilities.gstchecksum("07AAAAT7798M2ZW")
```
Output:
'Check Sum MISMATCH'


###4. Find out the last Correct Check sum for given 14 digit number

If we do not know the last digit of GST Number and want to find out, then we can use this function
to find out the correct last digit of the GST Code.

The Input should be at least 14 digit long. (pAssing even 15 digit will give the last correct digit)
```
from gst import gst_utilities

gst_utilities.getgstcheck("07AAAAT7798M2Z")
```

Output: K

```
from gst import gst_utilities

gst_utilities.getgstcheck("07AAAAT7798M2ZK")
```

Output: K

***Please note that this will only return whether the GST Number is Valid or Invalid.
It wont return the status of the GST Number like Suspended, Cancelled, Inactive***

Pro Tip For you : If you wish to check the Status of GST Number in Bulk, watch the below
video by Efficient Corporate [Bulk Check GST Numbers](https://www.youtube.com/watch?v=bGkvoky0X-M):)



###4. Extract PAN from GST Number

As we said above, the 3rd to 13th digit of GST Number is the PAN Number

So, simply do the below:

```
from gst import gst_utilities

gst_utilities.extract_pan("07AAAAT7798M2ZW")
```
Output: 'AAAAT7798M'


## License
Since, you have made the effort of reading the documentation till here, let me also explain in simple terms
what this license is all about.

Basically, this code is under a License GNU GPL, which basically means that you are free to use this code in your 
personal use or even use in your office.

And , interestingly, you can even give this code to someone else and also use this cde as a dependency in your own project

Preety much You can do everyting...But....

What you cannot do is to sell this code, or any of your project which uses this code with a commercial interest.

The Bottom Line is "Anything which you got for free, should be available freely..!!"

© 2021 Efficient Corporates

This repository is licensed under the OSI Approved :: GNU General Public License v3 or later (GPLv3+). See LICENSE file for details.


##For any issues / suggestions / complaints/ feedbacks / error faced / or even if you simply want to connect, we have our all ears for you...!! 

##JOIN our Community "EFFICIENT CORPORATES" NOW

[Youtube Channel](https://www.youtube.com/c/EFFICIENTCORPORATES):

[LinkedIn Company Page](https://www.linkedin.com/company/efficient-corporates/):

[LinkedIn Discussion Group](https://www.linkedin.com/groups/13967995/):

[Quora Space](https://efficientcorporates.quora.com/):

[Facebook Page](https://www.facebook.com/efficientcorporates):

[Twitter Handle](https://twitter.com/EfficientCorp):




