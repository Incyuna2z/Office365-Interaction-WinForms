# Office365-Interaction-WinForms

This sample list how to interact office365 in a WindowsFormApplication.

# How to re-create this sample

##prerequisites:
 1. Install Visual Studio 2017
 2. Azure Active Directory tenant
 3. Office365 Developer Subscription

##Steps 1: Create Project
 1. Open Visual Studio 2017
 2. Create a WindowsForms Application
 3. Add **System.Configuration** as reference
 
##Steps 2: Configure Office365
 1. In the Solution Explorer window,click **Connected Services -> Office 365 API**
 2. On the sign-in dialog box, enter the username and password for your Office 365 tenant
 3. After you're signed in, you will see a list of all the services.
 4. Initially, no permissions will be selected, as the app is not registered to consume any services yet.
 5. Select **My Files** dialog,select **Read your files**
 6. Click **OK**

After clicking OK in the Services Manager dialog box, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project.

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to App.config.(*The Office 365 API services use Azure AD to provide secure authentication to users' Office 365 data. To access the Office 365 APIs, you need to register your app with Azure AD*)

##Step3: Update ADAL version
Because we don't update ADAL version in Office365 provider yet,so we need to update it manually
 1. Open **Tools** -> **Nuget Package Manager** -> **Manage NugetPackage For Solution**
 2. Click **Update** tab, select **Microsoft.IndentityModel.Clients.ActiveDirctory**, update it to latest stable vesion

##Step4: Add AuthenticationHelp.cs and O365Help.cs
If you want to store temp token in local, you could define a new class which inherit from **TokenCache**

##Step5: Drag a button and ListView


