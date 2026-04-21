# GUI

`gui/PCOptimizer.Gui.ps1` is a WPF front-end for the existing `PC_Optimizer.ps1` engine.

- The GUI restarts itself as administrator
- The engine is launched with `-NonInteractive -NoRebootPrompt -EmitUiEvents`
- Progress is driven by UI events emitted from `PC_Optimizer.ps1`
- The generated HTML report can be opened directly from the GUI

Launch:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -STA -File .\gui\PCOptimizer.Gui.ps1
```
