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

#### Showng Supporting dialog
The Supporting may be show by calling this static method, depending on parameters and random check:
```csharp
public static void ShowIf(PluginControlBase plugin, bool manual, bool reload, AppInsights appins)
```

#### Showing current status by tool/user
It can be checked if this tool is enabled for Supporting by calling:
```csharp
public static bool IsEnabled(PluginControlBase plugin)
```
We can check if this tool is supported in this installation, and if so, by which type, by calling:
```csharp
public static SupportType IsSupporting(PluginControlBase plugin)
```

#### Example how to initialize Supporting
It might be called in the MyTool_Load event:
```csharp
Supporting.ShowIf(this, false, true, ai2);
if (Supporting.IsEnabled(this))
{
    tsbSupporting.Visible = true;
    var supptype = Supporting.IsSupporting(this);
    switch (supptype)
    {
        case SupportType.Company:
            tsbSupporting.Image = Resources.We_Support_icon;
            break;

        case SupportType.Personal:
            tsbSupporting.Image = Resources.I_Support_icon;
            break;

        case SupportType.Contribute:
            tsbSupporting.Image = Resources.I_Contribute_icon;
            break;
    }
}
else
{
    tsbSupporting.Visible = false;
}
```

### Updating

Update to the latest version of this repository; use this command in the tool repo:
```
git submodule update
```
*Note that either pull or push happens in Visual Studio - VS Code handles them more...*

## Requirements

The project needs to have a submodule to `https://github.com/rappen/Rappen.XTB.Helper` and added the project `Rappen.XTB.Helper` in the solution, and added it in the references in the tool project.
