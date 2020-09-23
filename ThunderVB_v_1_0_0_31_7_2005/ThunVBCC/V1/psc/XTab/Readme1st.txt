hi !!

thanx for downloading this control... :)


I'd enabled the subclassing in the code. If you want to debug the project without crashing VB, open the project, go to properties and set the IsDebug=0 flag to IsDebug=1. This will prevent subclassing. Once you are done with the control's testing/debugging. Open Properties again and set the IsDebug=1 to IsDebug=0. And then make a ocx. Once we make an ocx we don't get any GPFs while using it.


IMPORTANT: YOU MUST DISABLE THE SUBCLASSING IF YOU WANT TO DEBUG THE PROJECT. OTHERWISE VB will crash.

If you are not able to open the project due to some reason, mail me... maybe its due to subclassing. I'll tell you how to open the project without crashing VB.

my email address is :

nja91@yahoo.com
neeraj_agrawal_ind@rediffmail.com