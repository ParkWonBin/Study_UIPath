
```powershell
UiRobot <parameters>

Execution commands:
  -file <file_path> [-input <input_params>] [--rdp]
        -f,-file                   Workflow execution file
        -i,-input                  Dictionary of input parameters in JSON format
        --rdp                      Create interactive Windows session using RDP
  Examples:
    UiRobot -file "C:\UiPath\Project\Main.xaml"
    UiRobot -file "C:\UiPath\Project\Main.xaml" -input "{'inArg':'value'}" --rdp
    UiRobot -file "C:\UiPath\Project\project.json"
    UiRobot -file "C:\UiPath\Package\Notepad.1.0.6682.21636.nupkg"

Start process command:
  -process <process_name> [-input <input_params>]
        -p,-process                Process name (available version will be used)
        -i,-input                  Dictionary of input parameters in JSON format
  Examples:
    UiRobot -process UiPathDemo
    UiRobot -process "UiPathDemo" -input "{'inArg':'value'}"

Pack project command:
  -pack <project_path> -o <destination_folder> [-v <version>]
        -pack                      Project file path
        -o,-output                 Destination folder
        -v,-version                Package version
  Examples:
    UiRobot -pack "C:\UiPath\Project\project.json" -o "C:\UiPath\Package" -v 1.0.6820.22047

Connect to server commands:
  --connect [-url <server_url> -key <robot_key>] | [-connectionString <connection_string>]
  --disconnect
        --connect                  If used alone, it opens the Settings dialog
        --disconnect               Disconnect from server
        -url                       Set URL used to connect to server
        -key                       Set RobotKey used to connect to server
        -connectionString          Set the connection string for automatic deployment
  Examples:
    UiRobot --connect -url https://demo.uipath.com/ -key 696CCA0C-D347-48CE-8ADF-F65BBC2F15DE
    UiRobot --connect -connectionString https://demo.uipath.com/api/robotsservice/GetConnectionData?tenantId=1
    UiRobot --disconnect

Enable/Disable low level tracing:
  --enableLowLevel | --disableLowLevel
        --enableLowLevel           Enable low level tracing
        --disableLowLevel          Disable low level tracing

Licensing commands:
  --acquireLicense | --releaseLicense
        --acquireLicense           Acquire robot license from Orchestrator
        --releaseLicense           Release robot license from Orchestrator
```
