I attempted to install and run Azure CLI on my 64-bit Windows machine, however it fails to start with error 0xc000007b.
This error indicates a missing or mismatched Microsoft Visual C++ runtime dependency, which Azure CLI’s bundled Python requires.

To proceed, the system needs Microsoft Visual C++ 2015–2022 Redistributable (both x64 and x86) installed, or an IT-approved Azure CLI setup (preinstalled or via a managed dev environment).

Please advise on the approved approach.

Thank you.
