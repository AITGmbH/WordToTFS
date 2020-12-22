# WordToTFS

Since the beginning of TFS (Team Foundation Server), Microsoft has been providing integrations for Microsoft Office products like MS Excel or MS Project but not for MS Word.
WordToTFS has been created to fill this gap. Work items can be created and edited in MS Word and bidirectionally synced between TFS or Azure DevOps (Services and Server) and MS Word.
In addition to exporting and authoring work items, WordToTFS also supports generating test specifications and test result reports. More information about the functionality is available on the [website][ait-wordtotfs-page].

## Requirements & Installation

Install the following

1. MS Office (>= 2017))
2. Visual studio 2017 Community Edition (workloads: .NET Desktop Development, Office/SharePoint Development)
3. Visual Studio 2010 Tools for Office Runtime (>= version 60828.00, [download](https://www.microsoft.com/en-us/download/details.aspx?id=56961))
4. Windows SDK (xsd.exe is needed, see under C:\Program Files (x86)\Microsoft SDKs\Windows\)

Additional
  
* Add xsd.exe to Windows-Path (A post build is generating a W2T.xsd based on the class structure. The W2T is being used by the WordToTFS templates.)

## Local Build and Test

Build or Run TFS.SyncService.View.Word project.
The View.Word project is a COM plugin for Word.
When you run the solution with Visual Studio, Word will be opened and WordToTFS is directly loaded as COM-plugin.

## Automatic testing and unit testing

The project contains unit and integration tests.
Integration tests require a running Azure DevOps Server instance.
**For local testing, the Azure DevOps (Server) information must be provided in the Local.runsettings file.**
The tests are marked with an according test category:
  
* "None / No Class": Can be run without any Azure DevOps (Server).
* "ConnectionNeeded": A connection to any Azure DevOps (Server) is needed, but no preparations must be made within the project.
* "Interactive", "Interactive2" and "PreparationNeeded": These categories are currently not usable, because structured test data is needed. The test data is not yet part of this repository. For further details please see following paragraph.

The tests in the test category "Interactive", "Interactive2 and "PreparationNeeded" need already existing work items as predefined test data.
This setup is needed, to have certain states of work items which are used within the tests.
For that reason, we backed up a Project Collection containing those test data. This Project Collection is restored it in AITs test environment.
Since cleanup of the backup collection is much effort, it is not yet part of this repository.
A way of running these tests is an open point and may be shared in the future.
The backup contains test plans with test work items which are used to make end-to-end-test for word-file creation based on a configured template.
For example, this is needed in the class ConsoleExtensionHelperTest for testing the console behavior of WordToTFS.
The file ConfigurationNew1.xml contains the name and test result configuration or test specification configuration with the name of the test plan and test suite that will be used.
Please skip and ignore these tests currently. Somebody of the community might also work out a better test data management to automatically create the test data without using a backed up database.

## Centralized build system (Azure DevOps Pipeline) and Continuous Integration

As a centralized build system Azure DevOps Pipelines is used, because the previous builds were already using TFS-based build environments and all requirements could be fulfilled.
The continuous integration build is shared as a yaml workflow in this repository, please check the azure-pipelines-ci.yml.
To successfully run and configure the complete Azure DevOps Pipeline, a build environment with certain conditions and configurations must be set up.

The Pipeline needs the following variables configured:

* Test Server
  * TfsServerName: The name of the Azure DevOps (Server) used for testing
  * TfsServerFQDN: The fully qualified domain name within the network of the Azure DevOps (Server) used for testing
  * TfsTeamProjectCollectionUrl: The team project collection url to the Azure DevOps (Server) used for testing
  * TfsTeamProjectCollectionName: The team project collection name to the Azure DevOps (Server) used for testing
  * TeamProjectName: The team project name of the Azure DevOps (Server) used for testing
* Manifest update for having the project successfully configured for click once (updatemanifest_offline.ps1)
  * CertificateThumbPrint: The thumbprint of the pfx certificate that will be used
  * OfflinePublishUrl: The offline url where the click once result ist published
  * OnlinePublishUrl: The online url where the click once result ist published
  * TimestampUrl: The time stamp url for the certificate

## Strong Naming and Signing

The assemblies are strong named due to several restrictions using click once and office plugins.
A dummy.snk file is being used in all project files to strong name the assemblies and prevent dll manipulation.
For a publishing pipeline this publicly shared file should be replaced.

WordToTFS also needs to be signed via a certificate to successfully be used for publishing.
For local development the current configuration of TFS.SyncService.View.Word_TemporaryKey.pfx can be used. 
This pfx must be replaced and the valid ManifestCertificateThumbprint used for publishing.

## Documentation and User Guide

The full documation is currently only available on the [website][ait-wordtotfs-page].
The PDF contains detailed information about how to install, configure and use WordToTFS.

[ait-wordtotfs-page]: https://www.aitgmbh.de/blog/downloads/tfs-tools/ait-wordtotfs/

## Contribution

We love pull requests from everyone.
Fork, then clone the repo:

    git clone git@github.com:your-username/project-name.git

Make your change. Make sure to test your changes.
Push to your fork and please do not forget to submit a pull request.

At this point you're waiting on us. We like to at least comment on pull requests. We may suggest some changes or improvements or alternatives.

Some things that will increase the quality of your pull request:

* Write tests or test your changes.
* Apply the coding style of the project.
* Write a [good commit message][commit].
* Explain your changes.

[commit]: http://tbaggery.com/2008/04/19/a-note-about-git-commit-messages.html



