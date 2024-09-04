# Rappen.XTB.Supporting
## Help for XrmToolBox tool regarding supporting

The best way to use it is to att this repo in the tool solution with cmd:
```
git submodule add https://github.com/rappen/Rappen.XTB.Supporting
```

In the project, add existing item **As link**:
```
..\Rappen.XTB.Supporting\Forms\Supporting.cs
```

Now, THIS is the place where the code is updated, if needed, and available to all tools.

### Updating

Update to the latest version of this repository; use this command in the tool repo:
```
git submodule update
```
*Note that either pull or push happens in Visual Studio - VS Code handles them more...*

## Requirements

The project needs to have a submodule to `https://github.com/rappen/Rappen.XTB.Helper` and added the project `Rappen.XTB.Helper` in the solution, and added it in the references in the tool project.

## General Configs
There are several settings in the **[Config](Config)** folder.
