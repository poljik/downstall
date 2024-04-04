# Welcome

This is a script for automatic silent download and installation of the latest versions of applications.

## Examples

Examples with **-install** parameter:

```Powershell
./downstall.ps1 -install anydesk
./downstall.ps1 -install aimp, chrome, winrar
```

Or just:

```Powershell
./downstall.ps1 anydesk
./downstall.ps1 aimp, chrome, winrar
```

You can use **TAB** to display a list of options.
To download all applications without installing, you can ignore the arguments, for example:

```Powershell
./downstall.ps1
```

There are two more parameters **-downloadOnly**(without installation) and **-installOnly**(without download):

```Powershell
./downstall.ps1 -downloadOnly chrome, firefox
./downstall.ps1 -installOnly chrome, firefox
```
