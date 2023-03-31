$params = @{
    Path = "C:\Users\DylanHanssen\Documents\WindowsPowerShell\Modules\ManagementTool\ManagementTool.psd1"
    RootModule = "C:\Users\DylanHanssen\Documents\WindowsPowerShell\Modules\ManagementTool\ManagementTool.psm1"
    Author = "Dylan Hanssen"
    ModuleVersion = "0.9.0"
    Description = "Module with functions to reduce the Office workload"
    RequiredModule = "ImportExcel"
}
New-ModuleManifest @params

Test-ModuleManifest "C:\Users\DylanHanssen\Documents\WindowsPowerShell\Modules\ManagementTool\ManagementTool.psd1"