open System
open System.IO
open Microsoft.Win32
open System.Diagnostics
open System.Security.Principal
open System.Runtime.InteropServices

// ============================== Function Definitions ==============================

// ------------------------------ Win32 APIs ------------------------------

let deleteRegistryKey (hkey: RegistryHive) 
                 (subkey: string) 
                 (view: Option<RegistryView>) : Result<unit, Exception> =
    let reg_view = Option.defaultValue RegistryView.Default view
    try
        Ok(RegistryKey.OpenBaseKey(hkey, reg_view).DeleteSubKey(subkey))
    with
    | ex -> Error ex

// ------------------------------ Argument Parser ------------------------------

(* [!] Don't open `System.CommandLine` namespace because it conflicts with `FSharp.Core.Option` type. *)

type OfficeProgramEnum = 
    | Store = 10                // Office from MS Store
    | ClickToRun = 20           // Office Click-To-Run (for Office 15 ~ 16 (Office 2013 and later), install by ODT)
    | WindowsInstaller = 30     // Office MSI (for Office 2016 and earlier, install by mounting an setup image)

type ToolWorkModeEnum = 
    | Detect = 10
    | Remove = 20
    | Uninstall = 30

type OutputStateOptionEnum = 
    | Quiet = 0
    | Normal = 10
    | Verbose = 20

[<Struct>]
type ParsedArgs = {
    office_product_type: OfficeProgramEnum
    office_version: int
    work_mode: ToolWorkModeEnum
    keep_activation_info: bool
    no_restart: bool
    log_path: string
    output_state: OutputStateOptionEnum
    no_copyright_logo: bool
}

let printLogo (print_new_line: Option<bool>) = 
    printfn "Microsoft Office Remover"
    printfn "Github Repository: https://github.com/Wgpy66/MSOfficeRemover"
    printfn "Copyright (c) 2025 Wgpy66 All Rights Reserved."

    match print_new_line with
    | Some x -> 
        match x with
        | true -> printf "\n"
        | false -> ()
    | None -> ()

let mutable parse_result_struct: Option<ParsedArgs> = None

let parseArgument (args: string[]) = 
    // main command
    let command = CommandLine.RootCommand()
    // arguments
    let office_product_type_option = 
        let opt = CommandLine.Option<OfficeProgramEnum>("--office-product-type", "-p")
        opt.Description <- "Installed Office product type."
        opt.Required <- true
        opt
    let office_version_option =
        let opt = CommandLine.Option<int>("--office-version", "-o")
        opt.Description <- "Installed Office version."
        opt.Validators.Add(
            fun result ->
                let allowed_version = [| 12; 14; 15; 16 |]
                let inputed_version = result.GetValueOrDefault<int>()
                if not(Array.contains inputed_version allowed_version) then 
                    result.AddError $"Invalid Office version: {inputed_version}"
                else
                    ignore()
        )
        opt.Required <- true
        opt
    let work_mode_option = 
        let opt = CommandLine.Option<ToolWorkModeEnum>("--work-mode", "-m")
        opt.Description <- "Set work mode."
        opt.DefaultValueFactory <- 
            fun _ -> ToolWorkModeEnum.Detect
        opt
    let keep_activation_info_option =
        let opt = CommandLine.Option<bool>("--keep-activation-info", "-k")
        opt.Description <- 
            "Set whether keeping activation info is true. \n" + 
            "If you don't want to keep it, please do set this option."
        opt.Arity <- CommandLine.ArgumentArity.Zero
        opt
    let no_restart_option = 
        let opt = CommandLine.Option<bool>("--no-restart", "-r")
        opt.Description <- "Don't restart computer until removing is end."
        opt.Arity <- CommandLine.ArgumentArity.Zero
        opt
    let log_path_option = 
        let opt = CommandLine.Option<string>("--log-path", "-l")
        opt.Description <- "Set the path to store the log file."
        opt.DefaultValueFactory <- 
            fun _ -> "./log"
        opt
    let output_state_option = 
        let opt = CommandLine.Option<OutputStateOptionEnum>("--output-state", "-s")
        opt.Description <- "Output state."
        opt.DefaultValueFactory <- 
            fun _ -> OutputStateOptionEnum.Normal
        opt
    let no_copyright_logo_option =
        let opt = CommandLine.Option<bool>("--no-copyright-logo", "-c")
        opt.Description <- "Disabled to show copyright header."
        opt.Arity <- CommandLine.ArgumentArity.Zero
        opt
    (* let version_option = CommandLine.VersionOption("--version", "-v") *)
    // add arguments
    command.Add office_product_type_option
    command.Add office_version_option
    command.Add work_mode_option
    command.Add keep_activation_info_option
    command.Add no_restart_option
    command.Add log_path_option
    command.Add output_state_option
    command.Add no_copyright_logo_option
    (* command.Add version_option *)
    // add description
    command.Description <- "Microsoft Office removing commands."
    // add an action
    command.SetAction(
        fun pr -> 
            parse_result_struct <- Some (
                {
                    office_product_type = pr.GetValue office_product_type_option;
                    office_version = pr.GetValue office_version_option;
                    work_mode = pr.GetValue work_mode_option;
                    keep_activation_info = pr.GetValue keep_activation_info_option;
                    no_restart = pr.GetValue no_restart_option;
                    log_path = pr.GetValue log_path_option;
                    output_state = pr.GetValue output_state_option;
                    no_copyright_logo = pr.GetValue no_copyright_logo_option
                }
            )
            0
    )
    // parse args
    let parse_result = command.Parse args
    parse_result.Invoke()

// ------------------------------ Functions for Preparation ------------------------------
let self_path = Environment.ProcessPath

// check whether cuerrent administrator is running as privilege.
// if it is success, return true. otherwise, return false.
let checkAdminPriviledge () : bool =
    let administrator = WindowsBuiltInRole.Administrator
    let current_pricipal = WindowsPrincipal(WindowsIdentity.GetCurrent())
    current_pricipal.IsInRole administrator

// try getting the administraror priviledge
// if it is success, return `Some(())`, otherwise, return `None`.
let tryGetAdminPriviledge () : option<unit> = 
    let new_self_proc = ProcessStartInfo self_path
    new_self_proc.Arguments <- "--no-logo"
    new_self_proc.Verb <- "runas"
    Some ()

// ------------------------------ Program Entry Point ------------------------------

[<EntryPoint>]
let main (args: string[]) : int =
    printLogo (Some true)
    parseArgument args |> ignore
    0