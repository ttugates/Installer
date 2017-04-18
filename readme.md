# Installer

Installer will become a class library for my WinForm Projects.

# Current Features
- Fast, no fuss install
- Creates Desktop shortcut
- Creates Add/Remove Programs uninstall entry
- Silently, automatically stays updated when new releases are uploaded to GitHub
- Able to store and load Dlls/Assemblies in single EXE file as embeded resource

# Why
Ultimately this provides a simple solution for updating and maintiaining code I have written.

The process of releasing an update is as simple as:
- Opening the GitHub site and clicking create release. 
- Updating the Assembly version of the project in VS
- Building, dragging and dropping the exe to GitHub
- And editing the Tag to reflect the version.

# TODOs:
  - Address "this program may not have uninstalled properly" message
  - Extract this into a Shared Library, raising events instead of Message Boxes
  - Improve registry handling for combinations of x86 <=> x64 environment and build settings
  - Implement delete application exe on uninstall.
    - Options are: Via registry entry to run once on restart
    - Task scheduler?
    - Leave a minimal batch file
- Add a certificate / gain reputation with Microsoft Smart Screen Filter
