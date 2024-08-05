# Rappen.XTB.Supporting
## Help for XrmToolBox tool regarding supporting

Best way to use it is to att this repo in the tool solution with cmd:
```
git submodule add https://github.com/rappen/Rappen.XTB.Supporting
```

In the project, add existing item **As link**:
```
..\Rappen.XTB.Supporting\Forms\Supporting.cs
```

Now THIS is the place where the code is updated, if needed, and available too all tools.

### Updating

Update to latest version of this repository, use this command in the tool repo:
```
git submopdule update
```

## Requirements

The project needs to have a submodule to `https://github.com/rappen/Rappen.XTB.Helper` and added the project `Rappen.XTB.Helper` in the solution, and added it in the references in the tool project.