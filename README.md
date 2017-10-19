# spscGUI
Imports data from sharepoint/Sharpcloud to sharpcloud

### How to install from Visual Studio

Create a new C# WPF with the name spscGUI.

In the project folder, replace MainWindow.xaml and MainWindow.xaml.cs with the ones from this repo.

#Add References

Project -> Add References

System.Configuration

#Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Microsoft.Sharepoint.2013.Client.16


Install-Package SharpCloud.ClientAPI



###### WARNING

DEBUGGING THIS PROGRAM WILL TRIGGER ANTI VIRUS 