![Alt text](/SignatureSample.png?raw=true "Sample Signature")

# Meetings Today in Email Signature
A simple macro-driven technique to print timings of your meetings for the current day below your email signature. 

## Getting Started
Follow the instructions given below.

## Prerequisites
1. Microsoft Outlook (Macros should be enabled) 

## Installing
1. Save **2018Template.htm** to **C:\\users\\{yourusername}\\AppData\\Roaming\\Microsoft\\Signature**
2. Open Microsoft Outlook VBA (Press Alt+F11 from within Outlook) 
3. Paste contents of file **ThisOutlookSession.txt** to **ThisOutlookSession** module.
4. Import **modOutlookSignature.bas** as a new module.

## Enabling Macros
Navigate to ** Outlook -> Options -> Trust Center -> Trust Center Settings -> Macro Settings**. Select option _Notifications for all macros_.

## Selecting Signature
Navigate to Outlook -> Options -> Mail -> Signatures. Select **2018** as the signature for New message and Replies/forwards. If this is not available in the list, come back to this step later. 

## Testing
Open a new email. The _Meetings Today_ signature should appear as a additional text below your signature. If this doesn't appear, then perform the steps mentioned in under section _Selecting Signature_.  

## Licensing
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
