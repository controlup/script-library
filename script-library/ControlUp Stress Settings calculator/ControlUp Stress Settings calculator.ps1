[CmdletBinding()]
param (
    [parameter(Position=3, Mandatory = $false,
        HelpMessage = "Default Location of the folder with the ControlUp monitor export files.")]
    [string]$sourcepath = [Environment]::GetFolderPath("Desktop"),
    [parameter(Position=4, Mandatory = $false,
    HelpMessage = "Default location where csv's will be exported to.")]
    [string]$ExportPath = [Environment]::GetFolderPath("Desktop"),
    [parameter(Position=0,Mandatory = $false,
        HelpMessage = "Percentile used for yellow stress levels.")]
    [string]$yellowpercentile = 75,
    [parameter(Position=1,Mandatory = $false,
        HelpMessage = "Percentile used for red stress levels.")]
    [string]$redpercentile = 90,
    [parameter(Position=2,Mandatory = $false,
        HelpMessage = "Array with usernames or parts of usernames that shouldn't be taken into the calculations.")]
    [array]$filteredaccounts
)

<#
.SYNOPSIS
        This script can be used to calculate suggested Stress tresholds for ControlUp Real-Time Console
.DESCRIPTION
        This script uses either automated Monitor exports or Insights Exports as a base to calculate recommended Stres Tresholds for the ControlUp Console.
.PARAMETER Sourcepath
        This is an optional parameter pointing to a folder where the source files are located i.e. c:\dataset\
        Will default to the users desktop if the path can't be found or was not configured.
.PARAMETER Sourcepath
        This is an optional parameter pointing to a folder where the csv file will be exported to i.e. c:\export\
        Will default to the users desktop if the path can't be found or was not configured.
.PARAMETER yellowpercentile
        Number for the percentile to be used for recommending the Yellow Treshold
        Defaults to 75 if not configured.
.PARAMETER redpercentile
        Number for the percentile to be used for recommending the Yellow Treshold
        Defaults to 75 if not configured.
.PARAMETER Filteredaccounts
        Array of (parts of) accounts to be filtered out for the session metrics. i.e. "loginbot,wouterk,user1"
        No wildcards needed parts of accounts will also work.
.EXAMPLE
        If you want to specify the username and the session id:
        ./"Analyze Blast Bandwidth.ps1" "1" "controlup\samuel.legrand"
        In order to analyze the current session (no specific right needed):
        ./"Analyze Blast Bandwidth.ps1"
.OUTPUTS
        A list of the measured virtual channels with the bandwidth consumption in kbps.
.LINK
        https://www.powershellgallery.com/packages/Formulaic/0.2.1.0/Content/Get-Median.ps1
        https://gist.github.com/jbirley/f4c7775007aabbcf6b67b9160276b198
#>


if($verboseoutput -eq $True) {
    $VerbosePreference = "Continue"
}
if (test-path $sourcepath){
    $SourcePath=$sourcepath
    write-verbose -Message "Set default sourcepath to $sourcepath"
}
else{
    write-verbose -Message "Error validating sourcepath switching to default (Desktop)"
    [string]$SourcePath=[Environment]::GetFolderPath("Desktop")
}

if (test-path $ExportPath){
    $ExportPath=$ExportPath
    write-verbose -Message "Set default ExportPath to $ExportPath"
}
else{
    write-verbose -Message "Error validating exportpath switching to default (Desktop)"
    [string]$ExportPath=[Environment]::GetFolderPath("Desktop")
}

write-verbose -Message "Configured Yellow treshold percentile as $yellowpercentile "
write-verbose -Message "Configured Red treshold percentile as $redpercentile "

if($null -ne $filteredaccounts){
    write-verbose -message "Filtered accounts: $filteredaccounts."
    $Userfilter=New-Object System.Collections.ArrayList
    $filteredaccounts.Split(",") | foreach-object  {
        [void]$userfilter.add($_)
        write-verbose -message "Added $_ to user filter"
    }
}


$filteredmetrics="Id,Stress Level,AWS Owner,Hypervisor Platform,HZ Is Licensed,VM Snapshot Exists,VM Tools Version,Registered IP Addresses,NTP Status,Number of NICs,SSH enabled,vCPU/pCPU Ratio,GPU encoder Utilization,GPU Frame Buffer Memory Utilization,GPU decoder Utilization,Frames Per Second,Packet Loss"
$metricfilter=New-Object System.Collections.ArrayList
$filteredmetrics.Split(",") | foreach-object  {
    [void]$metricfilter.add($_)
    write-verbose -message "Added $_ to Metric filter"
}

# Create arrays for both monitor and Insights related data
$Monitor_Sourcedata=New-Object System.Collections.ArrayList
$Monitor_Data=New-Object System.Collections.ArrayList
$Insights_Sourcedata=New-Object System.Collections.ArrayList
$Insights_Data=New-Object System.Collections.ArrayList

# Load the assemblies for the GUI
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
add-Type -AssemblyName WindowsBase
#Region XAML
$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ControlUp Stress Settings Calculator" Height="550" Width="1024">
    <Grid>
        <TabControl>
            <TabItem Header="Monitor Data">
                <ScrollViewer
                    HorizontalScrollBarVisibility = "auto"
                    VerticalScrollBarVisibility = "auto">
                    <Canvas Height="440" Width="950">
                        <TextBox
                            x:Name = "Monitor_Folder_TextBox"
                            HorizontalAlignment="Left"
                            VerticalAlignment="Top"
                            Height="auto"
                            Width="200"
                            Margin="20,20,0,0"
                            Text="$SourcePath"
                            HorizontalScrollBarVisibility = "auto"
                        />
                        <Button x:Name="Monitor_Select_Folder"
                            Content="Browse"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,20,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                            ToolTip="See the Help tab for more information on how to get the required files from the ControlUp Monitors."
                        />
                        <Button x:Name="Monitor_Load_Data"
                            Content="Load Data"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="380,20,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                            ToolTip="Depending on the amount of files and file sizes it might take some time to load the data."
                        />
                        <ComboBox x:Name="Monitor_SourceType"
                            IsEditable="True"
                            IsReadOnly="True"
                            Text = "Select Data Source"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="200"
                            Margin="20,70,0,0"
                            VerticalAlignment="Top"
                            IsSynchronizedWithCurrentItem="True"
                        >
                                <ComboBoxItem Content="Sessions"/>
                                <ComboBoxItem Content="Hosts"/>
                                <ComboBoxItem Content="Machines"/>
                                <ComboBoxItem Content="Folders"/>
                        </ComboBox>
                        <ComboBox x:Name="Monitor_Available_Metrics"
                            IsEditable="True"
                            IsReadOnly="True"
                            Text = "Load Data first"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="200"
                            Margin="20,100,0,0"
                            VerticalAlignment="Top"
                            ToolTip="Load data first before metrics are visible"
                        />
                        <ComboBox x:Name="Monitor_Folder_Selector"
                            IsEditable="True"
                            IsReadOnly="True"
                            Text = "Select a folder"
                            HorizontalAlignment="Left"
                            Visibility = "Hidden"
                            Height="20"
                            Width="200"
                            Margin="20,130,0,0"
                            VerticalAlignment="Top"
                            ToolTip="{Binding RelativeSource={RelativeSource Self}, Path=SelectedValue}"
                        >
                        </ComboBox>
                        <DataGrid x:Name="Monitor_Datagrid" AutoGenerateColumns="True"
                            Margin="20,160,0,0"
                            Height="200"
                            Width="800"
                            IsReadOnly = "true"
                            HorizontalScrollBarVisibility = "Auto"
                            VerticalScrollBarVisibility = "Auto"
                        >
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Metric Name" Binding="{Binding Name}" Width="auto" >
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding name}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Metric Source" Binding="{Binding Source}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Source}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Source Folder" Binding="{Binding Folder}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Folder}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Suggested Stress Treshold" Binding="{Binding Stress_Level}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Stress_Level}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Calculation" Binding="{Binding Calculation}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Calculation}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Value}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Datapoints" Binding="{Binding Datapoints}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Datapoints}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                            </DataGrid.Columns>
                        </DataGrid>
                        <Button x:Name="Monitor_Export"
                            Content="Export to CSV"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="20,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <Button x:Name="Monitor_Copy"
                            Content="Copy CSV Data"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="140,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <Button x:Name="Monitor_Clear"
                            Content="Clear"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="260,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <CheckBox x:Name="Monitor_Show_Median"
                            Content = 'Show Median'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,70,0,0"
                            VerticalAlignment="Top"
                        />
                        <CheckBox x:Name="Monitor_Show_Average"
                            Content = 'Show Average'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,100,0,0"
                            VerticalAlignment="Top"
                        />
                        <CheckBox x:Name="Monitor_Ignore_0_values"
                            Content = 'Ignore 0 values'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="150"
                            Margin="380,70,0,0"
                            VerticalAlignment="Top"
                            ToolTip="If checked values that have a value of 0 or that show N/A in the grid will be ignored"
                        />
                        <Image x:Name="Monitor_Image" Height="150" Canvas.Left="510" Canvas.Top="5" Width="112"/>
                            <TextBlock
                            Height="20"
                            Width="20"
                            HorizontalAlignment="Left"
                            VerticalAlignment="Top"
                            TextWrapping="Wrap"
                            Margin="540,50,0,0"
                            ToolTip="Isn't VEronica awesome?"
                        />
                    </Canvas>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Insights Data">
                <ScrollViewer
                    HorizontalScrollBarVisibility = "auto"
                    VerticalScrollBarVisibility = "auto">
                    <Canvas Height="440" Width="950">
                        <TextBox
                            x:Name = "Insights_File_TextBox"
                            HorizontalAlignment="Left"
                            VerticalAlignment="Top"
                            Height="auto"
                            Width="200"
                            Margin="20,20,0,0"
                            Text="$SourcePath"
                            ToolTip="$sourcepath"
                            HorizontalScrollBarVisibility = "auto"
                        />
                        <Button x:Name="Insights_Select_File"
                            Content="Browse"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,20,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                            ToolTip="See the Help tab for more information on how to get the required files from ControlUp Insights."
                        />
                        <ComboBox x:Name="Insights_Available_Metrics"
                            IsEditable="True"
                            IsReadOnly="True"
                            Text = "Load Data first"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="200"
                            Margin="20,100,0,0"
                            VerticalAlignment="Top"
                            ToolTip="Load data first before metrics are visible"
                        />
                        <ComboBox x:Name="Insights_Folder_Selector"
                            IsEditable="True"
                            IsReadOnly="True"
                            Text = "Select a folder"
                            HorizontalAlignment="Left"
                            Visibility = "Hidden"
                            Height="20"
                            Width="200"
                            Margin="20,130,0,0"
                            VerticalAlignment="Top"
                            ToolTip="{Binding RelativeSource={RelativeSource Self}, Path=SelectedValue}"
                        >
                        </ComboBox>
                        <DataGrid x:Name="Insights_Datagrid" AutoGenerateColumns="True"
                            Margin="20,160,0,0"
                            Height="200"
                            Width="580"
                            IsReadOnly = "true"
                            HorizontalScrollBarVisibility = "Auto"
                            VerticalScrollBarVisibility = "Auto"
                        >
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Metric Name" Binding="{Binding Name}" Width="auto" >
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding name}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                    </DataGridTextColumn>
                                <DataGridTextColumn Header="Suggested Stress Treshold" Binding="{Binding Stress_Level}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                    <Style TargetType="DataGridCell">
                                        <Setter Property="ToolTip" Value="{Binding Stress_Level}" />
                                    </Style>
                                </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Calculation" Binding="{Binding Calculation}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Calculation}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="auto">
                                    <DataGridTextColumn.CellStyle>
                                        <Style TargetType="DataGridCell">
                                            <Setter Property="ToolTip" Value="{Binding Value}" />
                                        </Style>
                                    </DataGridTextColumn.CellStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Datapoints" Binding="{Binding Datapoints}" Width="auto">
                                <DataGridTextColumn.CellStyle>
                                    <Style TargetType="DataGridCell">
                                        <Setter Property="ToolTip" Value="{Binding Datapoints}" />
                                    </Style>
                                </DataGridTextColumn.CellStyle>
                            </DataGridTextColumn>
                            </DataGrid.Columns>
                        </DataGrid>
                        <Button x:Name="Insights_Export"
                            Content="Export to CSV"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="20,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <Button x:Name="Insights_Load_Data"
                            Content="Load Data"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="380,20,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                            ToolTip="Depending on the file size it might take some time to load the data."
                        />
                        <Button x:Name="Insights_Copy"
                            Content="Copy CSV Data"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="140,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <Button x:Name="Insights_Clear"
                            Content="Clear"
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="260,390,0,0"
                            VerticalAlignment="Top"
                            ClickMode="Press"
                        />
                        <CheckBox x:Name="Insights_Show_Median"
                            Content = 'Show Median'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,70,0,0"
                            VerticalAlignment="Top"
                        />
                        <CheckBox x:Name="Insights_Show_Average"
                            Content = 'Show Average'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="100"
                            Margin="250,100,0,0"
                            VerticalAlignment="Top"
                        />
                        <CheckBox x:Name="Insights_Ignore_0_values"
                            Content = 'Ignore 0 values'
                            HorizontalAlignment="Left"
                            Height="20"
                            Width="150"
                            Margin="380,70,0,0"
                            VerticalAlignment="Top"
                            ToolTip="If checked values that have a value of 0 will be ignored"
                        />
                        <Image x:Name="Insights_Image" Height="150" Canvas.Left="510" Canvas.Top="5" Width="112"/>
                        <TextBlock
                            Height="20"
                            Width="20"
                            HorizontalAlignment="Left"
                            VerticalAlignment="Top"
                            TextWrapping="Wrap"
                            Margin="540,50,0,0"
                            ToolTip="Isn't VEronica awesome?"
                        />
                    </Canvas>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Help">
                <ScrollViewer
                    HorizontalScrollBarVisibility = "auto"
                    VerticalScrollBarVisibility = "auto">
                    <Canvas Height="440" Width="900">
                        <RichTextBox
                            HorizontalScrollBarVisibility = "Disabled"
                            VerticalScrollBarVisibility = "auto"
                            Height="65"
                            Width="880"
                            BorderThickness="0"
                            IsReadOnly="True"
                            Margin="10,10,0,0">
                            <FlowDocument>
                                <Paragraph  Margin="6" TextAlignment="Left" FontWeight="Bold" FontSize="18">About</Paragraph>
                                <Paragraph  Margin="6" TextAlignment="Left">The ControlUp Stress Settings Calculator enables you to use your own data to determine how to set your Stress Thresholds in the Real-Time Console.
                                <LineBreak/>
                                For a full guide on using the tool, access the Knowledge Base article linked from the bottom of this page.
                                </Paragraph>
                            </FlowDocument>
                        </RichTextBox>
                            <RichTextBox
                                HorizontalScrollBarVisibility = "Disabled"
                                VerticalScrollBarVisibility = "auto"
                                Height="30"
                                Width="880"
                                BorderThickness="0"
                                IsReadOnly="True"
                                Margin="10,75,0,0">
                                <FlowDocument>
                                    <Paragraph Margin="6" TextAlignment="Left" FontWeight="Bold" FontSize="18">Usage</Paragraph>
                                </FlowDocument>
                            </RichTextBox>
                            <RichTextBox
                                HorizontalScrollBarVisibility = "Disabled"
                                VerticalScrollBarVisibility = "auto"
                                Height="280"
                                Width="880"
                                BorderThickness="0"
                                IsReadOnly="True"
                                Margin="10,100,0,0">
                                <FlowDocument>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Browse</Bold>
                                    <LineBreak/>
                                    Select the location of the CSV file exported from Insights data, or the folder for the exported Monitor Data. Alternatively, copy/paste the location into the empty field.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Load Data</Bold>
                                    <LineBreak/>
                                    Once you selected or entered the location, click to load the data into the tool.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Select Data Source</Bold>
                                    <LineBreak/>
                                    For Monitor Data only, select from the dropdown options, which include: Sessions, Hosts, Machines, Folders - per the views displayed in the Real-Time Console.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Select Metric</Bold>
                                    <LineBreak/>
                                    Use the dropdown list to select which metric you want to see results for and set stress thresholds.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left" FontWeight="Bold" FontSize="18">Options</Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Select a folder</Bold>
                                    <LineBreak/>
                                    When you select Folders, Hosts or Machines as the Data Source for Monitor Data, you can filter the metrics to calculate by selecting a folder from your organization tree.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Show Median</Bold>
                                    <LineBreak/>
                                    Select to also display the median value of the imported metrics.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Show average</Bold>
                                    <LineBreak/>
                                    Select to also display the average value of the imported metrics.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left"><Bold>Ignore 0 values</Bold>
                                    <LineBreak/>
                                    Select to ignore all values of 0 or n/a that might otherwise skew the threshold percentage values.
                                    </Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left" FontWeight="Bold" FontSize="18">Tips</Paragraph>
                                    <Paragraph Margin="6" TextAlignment="Left">- Once the data comes in, you can drag the column widths to easily view the values.
                                    <LineBreak/>
                                    - The <Bold>Suggested Stress Threshold</Bold> column gives you the stress level you should define for Yellow and Red values.
                                    <LineBreak/>
                                    - If you will need to refer to this data, you can select <Bold>Export to CSV</Bold> or <Bold>Copy CSV Data</Bold>.
                                    <LineBreak/>
                                    - Click <Bold>Clear</Bold> to start a new calculation.
                                    </Paragraph>
                                </FlowDocument>
                        </RichTextBox>
                            <RichTextBox
                            HorizontalScrollBarVisibility = "Disabled"
                            VerticalScrollBarVisibility = "auto"
                            Height="30"
                            Width="880"
                            BorderThickness="0"
                            IsReadOnly="True"
                            Margin="10,380,0,0">
                            <FlowDocument>
                                <Paragraph Margin="6" TextAlignment="Left" FontWeight="Bold" FontSize="18">Links</Paragraph>
                            </FlowDocument>
                        </RichTextBox>
                        <Label x:Name="Help_Link1" Content="Full instructions in this Knowledge Base article." HorizontalAlignment="Left" Margin="16,400,0,0" VerticalAlignment="Top" Width="auto" FontSize='14' Foreground='DarkBlue' Cursor="Hand" ToolTip='ControlUp KB'/>
                        <Label x:Name="Help_Link2" Content="Video guide on YouTube." HorizontalAlignment="Left" Margin="16,415,0,0" VerticalAlignment="Top" Width="auto" FontSize='14' Foreground='DarkBlue' Cursor="Hand" ToolTip='Youtube'/>
                    </Canvas>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
#Endregion

#region Generic Functions

function DecodeBase64Image {
    param (
        [Parameter(Mandatory=$true)]
        [String]$ImageBase64
    )
    # Parameter help description
    $ObjBitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage #Provides a specialized BitmapSource that is optimized for loading images using Extensible Application Markup Language (XAML).
    $ObjBitmapImage.BeginInit() #Signals the start of the BitmapImage initialization.
    $ObjBitmapImage.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($ImageBase64) #Creates a stream whose backing store is memory.
    $ObjBitmapImage.EndInit() #Signals the end of the BitmapImage initialization.
    $ObjBitmapImage.Freeze() #Makes the current object unmodifiable and sets its IsFrozen property to true.
    $ObjBitmapImage
}
Function Get-filecontent {
    # Uses iostream to read the csv files
    Param(
    [string]$file
    )
    Process
    {
        $read = New-Object System.IO.StreamReader($file)
        #$filecontent = @()
        $filecontent=New-Object System.Collections.Generic.List[System.String]

        while (($line = $read.ReadLine()) -ne $null)
        {
            $filecontent.add($line)
        }

        $read.Dispose()
        return $filecontent
    }
}

# Function sourced from https://gist.github.com/jbirley/f4c7775007aabbcf6b67b9160276b198
function Get-Percentile {

    <#
    .SYNOPSIS
        Returns the specified percentile value for a given set of numbers.

    .DESCRIPTION
        This function expects a set of numbers passed as an array to the 'Sequence' parameter.  For a given percentile, passed as the 'Percentile' argument,
        it returns the calculated percentile value for the set of numbers.

    .PARAMETER Sequence
        A array of integer and/or decimal values the function uses as the data set.
    .PARAMETER Percentile
        The target percentile to be used by the function's algorithm.

    .EXAMPLE
        $values = 98.2,96.5,92.0,97.8,100,95.6,93.3
        Get-Percentile -Sequence $values -Percentile 0.95

    .NOTES
        Author:  Jim Birley
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Double[]]$Sequence
        ,
        [Parameter(Mandatory)]
        [Double]$Percentile
    )

    $Sequence = $Sequence | Sort-Object
    [int]$N = $Sequence.Length
    #Write-Verbose "N is $N"
    [Double]$Num = ($N - 1) * $Percentile + 1
    #Write-Verbose "Num is $Num"
    if ($num -eq 1) {
        return $Sequence[0]
    } elseif ($num -eq $N) {
        return $Sequence[$N-1]
    } else {
        $k = [Math]::Floor($Num)
        [Double]$d = $num - $k
        return $Sequence[$k - 1] + $d * ($Sequence[$k] - $Sequence[$k - 1])
    }
}


# sourced from https://www.powershellgallery.com/packages/Formulaic/0.2.1.0/Content/Get-Median.ps1
function Get-Median{
    <#
    .Synopsis
        Gets a median
    .Description
        Gets the median of a series of numbers
    .Example
        Get-Median 2,4,6,8
    .Link
        Get-Average
    .Link
        Get-StandardDeviation
    #>
    param(
    # The numbers to average
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
    [Double[]]
    $Number
    )

    begin {
        $numberSeries = @()
    }

    process {
        $numberSeries += $number
    }

    end {
        $sortedNumbers = @($numberSeries | Sort-Object)
        if ($numberSeries.Count % 2) {
            # Odd, pick the middle
            $sortedNumbers[($sortedNumbers.Count / 2) - 1]
        } else {
            # Even, average the middle two
            ($sortedNumbers[($sortedNumbers.Count / 2)] + $sortedNumbers[($sortedNumbers.Count / 2) - 1]) / 2
        }
    }
}


function Is-Numeric ($Value) {
    # Checks if a value is numeric
    return $Value -match "^[\d\.]+$"
}

#endregion

#region Monitor calculator functions

function monitor-clear_selected_metric{
    $dontcalculate=$true
    $Monitor_Available_Metrics.Text="Select Metric."
    $Monitor_Available_Metrics.ToolTip="Select Metric."
}

Function Clear-Monitor_data{
    $Monitor_Datagrid.items.clear()
    $dontcalculate=$true
    $Monitor_Available_Metrics.Text="Select Metric."
    $Monitor_Available_Metrics.ToolTip="Select Metric."
}

Function copy-Monitor_data{
    # Copies the data of the monitor datagrid to the clipboard
    Set-Clipboard -value ($Monitor_Datagrid.items | convertto-csv -NoTypeInformation)
}

Function export-Monitor_data{
    # Exports data on the monitor datagrid to CSV
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $ExportPath
    $SaveFileDialog.filter = "csv (*.csv)| *.csv"
    $SaveFileDialog.filename="Monitor Stress Levels Calculation.csv"
    $Show = $SaveFileDialog.ShowDialog()
    If ($Show -eq "OK")
    {
        $Monitor_Datagrid.items | Export-Csv $SaveFileDialog.filename
    }
    Else
    {
    }
}

function set-Monitor_folderfilter {
    # Makes it possible to filter on folder name if applicable
    $filetype=$Monitor_SourceType.SelectedItem.Content
    $Monitor_Data.clear()
    $folderchoice=$Monitor_Folder_Selector.SelectedItem
    if ($folderchoice -eq "All Folders"){
        $Monitor_Sourcedata.foreach{[void]$Monitor_Data.add($_)}
    }
    elseif ($filetype -eq "Hosts" -OR $filetype -eq "Folders"){
        ($Monitor_Sourcedata.where{$_."Folder path" -eq $folderchoice}).foreach{$Monitor_Data.add($_)}
    }
    elseif ($filetype -eq "Machines"){
        ($Monitor_Sourcedata.where{$_."Folder" -eq $folderchoice}).foreach{$Monitor_Data.add($_)}
        }
    $dontcalculate=$true
    $Monitor_Available_Metrics.Text="Select Metric."
    $Monitor_Available_Metrics.ToolTip="Select Metric."
}

function get-Monitor_stats{
    #Calculates the values of the selected metric
    $useforyellow="0."+$yellowpercentile
    $useforred="0."+$redpercentile
    $Monitor_Available_Metrics.ToolTip=$Monitor_Available_Metrics.SelectedItem
    $metricname=$Monitor_Available_Metrics.SelectedItem
    $filetype=$Monitor_SourceType.SelectedItem.Content
    $folderchoice=$Monitor_Folder_Selector.SelectedItem
    if($NULL -eq $folderchoice -OR $folderchoice -eq "All Folders"){
        $folderchoice="n/a"
    }

    if(!$metricname -AND $Monitor_Available_Metrics.Text -ne "Disabled while loading data" -AND $Monitor_Available_Metrics.Text -ne "Select Metric.") {
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Select a Metric first"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    elseif($Monitor_Available_Metrics.Text -eq "Disabled while loading data"){
        # Just need to not run get stats
    }
    elseif($dontcalculate -eq $true){
        $dontcalculate=$false
    }
    else{
        try {
            if($Monitor_Ignore_0_values.IsChecked -eq "True"){
                $filtereddatapoints=@()
                $filtereddatapoints=$Monitor_Data.where{$_.$metricname -ne "0" -AND $_.$metricname -ne "0.00" -AND $_.$metricname -ne "0,00"}
                #write-verbose -message $filtereddatapoints.$metricname
                $yellow=Get-Percentile -Sequence ($filtereddatapoints.$metricname) -Percentile $useforyellow
                $red=Get-Percentile -Sequence ($filtereddatapoints.$metricname) -Percentile $useforred
                $average=[Linq.Enumerable]::average([decimal[]] @($filtereddatapoints.$metricname))
                $median=(get-median $filtereddatapoints.$metricname)
                $Monitor_Datapoints=$filtereddatapoints.count
            }
            else{
                $yellow=Get-Percentile -Sequence ($Monitor_Data.$metricname) -Percentile $useforyellow
                $red=Get-Percentile -Sequence ($Monitor_Data.$metricname) -Percentile $useforred
                $average=[Linq.Enumerable]::average([decimal[]] @($Monitor_Data.$metricname))
                $median=(get-median $Monitor_Data.$metricname)
                $Monitor_Datapoints=$Monitor_Data.count
            }
            [void]$Monitor_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Source=$filetype;Calculation=$yellowpercentile+"th Percentile";Value=[math]::Round($yellow,2);Datapoints=$Monitor_Datapoints;Folder=$folderchoice;Stress_level="Change Yellow to "+[math]::Round($yellow)})
            [void]$Monitor_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Source=$filetype;Calculation=$redpercentile+"th percentile";Value=[math]::Round($red,2);Datapoints=$Monitor_Datapoints;Folder=$folderchoice;Stress_level="Change Red to "+[math]::Round($red)})
            if($Monitor_Show_Average.IsChecked -eq "True"){
                [void]$Monitor_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Source=$filetype;Calculation="Average";Value=[math]::Round($Average,2);Datapoints=$Monitor_Datapoints;Folder=$folderchoice})
            }
            if($Monitor_Show_Median.IsChecked -eq "True"){
                [void]$Monitor_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Source=$filetype;Calculation="Median";Value=[math]::Round($Median,2);Datapoints=$Monitor_Datapoints;Folder=$folderchoice})
            }
        }
        catch{
            write-verbose -message "Calculation failed this metric probably couldn't be converted to an integer"
        }
    }
}

function load-Monitor_data {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    # Loads the data for the monitor export files
    $Monitor_Sourcedata.clear()
    $Monitor_Data.clear()
    $Monitor_Folder_Selector.items.clear()
    $Monitor_Available_Metrics.Text="Disabled while loading data"
    $Monitor_Available_Metrics.items.clear()
    $sourcefolder=$Monitor_Folder_TextBox.Text
    $filetype=$Monitor_SourceType.SelectedItem.Content
    write-verbose -message "Gathering metrics for $filetype"
    if ($filetype -eq "Use Browse or enter the source file location"){
        $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
        $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
        $MessageBody = "Please select a metric type first"
        $MessageTitle = "Warning"
        [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    elseif($sourcefolder -eq ""){
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Select a source path first"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    elseif($filetype -eq "" -OR $Null -eq $filetype){
        $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
        $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
        $MessageBody = "Select a data source first"
        $MessageTitle = "Warning"
        [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    elseif(!(test-path $sourcefolder)){
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Path not found"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    else {
        $files = (get-childitem $sourcefolder).where({$_.name -like "*$filetype*"})
        $filecount=$files.count
        write-verbose -message "importing $filecount files"
        if($files.count -gt 0){
            $header=((get-childitem $sourcefolder).where({$_.name -like "*$filetype*"}) | select-object -first 1 | get-content | Select-Object -skip 1 -first 1).Split(",")
            foreach ($file in $files){
                $content=$null
                if ($filetype -eq "sessions"){
                    [array]$content=((Get-filecontent -file $file.fullname).replace('","','"^"') | select-object -skip 2 | out-string | ConvertFrom-Csv -Header $header -Delimiter ^).where({$_.user -ne ""})
                    if ($Userfilter){
                        [array]$content=$content.where({$_.user |select-string -pattern $userfilter -notmatch })
                    }
                }
                else {
                    [array]$content=((Get-filecontent -file $file.fullname).replace('","','"^"') | select-object -skip 2 | out-string | ConvertFrom-Csv -Header $header -Delimiter ^)
                }
                $content.foreach({[void]$Monitor_Sourcedata.add($_)})
            }
            $stopwatch.stop()
            $elapsedtime=$stopwatch.Elapsed.TotalSeconds
            write-verbose -message "reading the files took $elapsedtime seconds"
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $header=$header.where({$_ |select-string -pattern $metricfilter -notmatch }) 
            $header=$header | Sort-Object
            foreach ($i in $header){
                $p=$Monitor_Sourcedata."$i" | where-object {$_ -ne ""} | select-object -first 1
                try{
                    get-median $p | out-null
                    Get-Percentile -sequence $p -Percentile 0.75 | out-null
                    [void] $Monitor_Available_Metrics.Items.Add($i)
                }
                catch{}
            }
            $stopwatch.stop()
            $elapsedtime=$stopwatch.Elapsed.TotalSeconds
            write-verbose -message "Checking the metrics took $elapsedtime seconds"
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            if ($filetype -eq "hosts" -OR $filetype -eq "Folders"){
                $Monitor_Folder_Selector.Visibility="Visible"
                $foldermenuitems= $Monitor_Sourcedata."Folder path" | select-object -unique
                $foldermenuitems | sort-object | foreach-object {[void] $Monitor_Folder_Selector.Items.Add($_)}
                [void] $Monitor_Folder_Selector.Items.Add("All Folders")
            }
            elseif ($filetype -eq "Machines"){
                $Monitor_Folder_Selector.Visibility="Visible"
                $foldermenuitems= $Monitor_SourceData."Folder" | select-object -unique
                $foldermenuitems | sort-object | foreach-object {[void] $Monitor_Folder_Selector.Items.Add($_)}
                [void] $Monitor_Folder_Selector.Items.Add("All Folders")
            }
            else{
                $Monitor_Folder_Selector.Visibility="Hidden"
            }
            $Monitor_SourceData.foreach{[void]$Monitor_Data.add($_)}
            $Monitor_Available_Metrics.Text="Select Metric"
            $Monitor_Available_Metrics.ToolTip="Select Metric"
        }
        else{
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Could not find any suitable files in this location"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        }
    }
    $stopwatch.stop()
    $elapsedtime=$stopwatch.Elapsed.TotalSeconds
    write-verbose -message "Building the menu took $elapsedtime seconds"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $count = $Monitor_Data.count
    write-verbose -message "Object count: $count"
}

function select-Monitor_inputtype {
    # Sets the tooltip of the selected metric to show that metric
    $Monitor_SourceType.ToolTip=$Monitor_SourceType.SelectedItem.Content
}

function Select-Monitor_FolderDialog {
    # Allows for selecting the folder that holds the csv files
    [string]$Description="Select Folder"
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.SelectedPath = $SourcePath
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK")
    {
        $Monitor_Folder_TextBox.Text=$objForm.SelectedPath
    }
    Else
    {
    }
}
#endregion

#region Insights related functions

function insights-clear_selected_metric{
    $dontcalculate=$true
    $Insights_Available_Metrics.Text="Select Metric."
    $Insights_Available_Metrics.ToolTip="Select Metric."
}

function open-Insights_FileDialog {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.OpenFileDialog
    $objForm.InitialDirectory  = $SourcePath
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK")
    {
        $Insights_File_TextBox.Text=$objForm.FileName
        $Insights_File_TextBox.Tooltip=$objForm.FileName
    }
    Else
    {
        write-verbose -message "Operation cancelled by user."
    }
}

Function Clear-Insights_data{
    $Insights_Datagrid.items.clear()
    $dontcalculate=$true
    $Insights_Available_Metrics.Text="Select Metric."
    $Insights_Available_Metrics.ToolTip="Select Metric."
    }

Function copy-Insights_data{
    Set-Clipboard -value ($Insights_datagrid.items | convertto-csv -NoTypeInformation)
    }

Function export-Insights_data{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $ExportPath
    $SaveFileDialog.filter = "csv (*.csv)| *.csv"
    $SaveFileDialog.filename="Insights Stress levels Calculation.csv"

    $Show = $SaveFileDialog.ShowDialog()
    If ($Show -eq "OK")
    {
        $Insights_datagrid.items | Export-Csv $SaveFileDialog.filename
    }
    Else
    {
        write-verbose -message "Operation cancelled by user."
    }
}
function load-Insights_data {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Insights_sourcedata.clear()
    $Insights_data.clear()
    $Insights_Available_Metrics.Text="Disabled while loading data"
    $Insights_Available_Metrics.items.clear()
    $sourcefile=$Insights_File_TextBox.Text
    if(!(test-path $sourcefile)){
            write-verbose -message "File $sourcefile not found."
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "File not found"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    else {
        write-verbose -message "Reading $sourcefile."
        ((Get-filecontent -file $sourcefile | convertfrom-csv)).foreach{[void]$Insights_sourcedata.add($_)}
        if ($Userfilter){
            $Insights_sourcedata=$Insights_sourcedata | where-object {$_."user name" |select-string -pattern $userfilter -notmatch }
        }
        $header=$Insights_sourcedata | get-member -type noteproperty | select-object -expandproperty name
    }
    $stopwatch.stop()
    $elapsedtime=$stopwatch.Elapsed.TotalSeconds
    write-verbose -message "reading the file took $elapsedtime seconds"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($i in $header){
        $p=$Insights_sourcedata."$i" | where-object {$_ -ne ""} | select-object -first 1
        try{
            get-median $p | out-null
            Get-Percentile -sequence $p -Percentile 0.75 | out-null
            [void] $Insights_Available_Metrics.Items.Add($i)
        }
        catch{}
    }
    $stopwatch.stop()
    $elapsedtime=$stopwatch.Elapsed.TotalSeconds
    write-verbose -message "Checking the metrics took $elapsedtime seconds"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Insights_sourcedata.foreach{[void]$Insights_data.add($_)}
    $Insights_Available_Metrics.Text="Select Metric"
    $Insights_Available_Metrics.ToolTip="Select Metric"
    $stopwatch.stop()
    $elapsedtime=$stopwatch.Elapsed.TotalSeconds
    write-verbose -message "Building the menu took $elapsedtime seconds"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $count = $Insights_data.count
    write-verbose -message "Object count: $count"
}

function get-Insights_stats{
    $useforyellow="0."+$yellowpercentile
    $useforred="0."+$redpercentile
    $Insights_Available_Metrics.ToolTip=$Insights_Available_Metrics.SelectedItem
    $metricname=$Insights_Available_Metrics.SelectedItem
    if(!$metricname -AND $Insights_Available_Metrics.Text -ne "Disabled while loading data" -AND $Insights_Available_Metrics.Text -ne "Select Metric.") {
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Select a Metric first"
            $MessageTitle = "Warning"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    }
    elseif($Insights_Available_Metrics.Text -eq "Disabled while loading data"){
        # Don't need to do nothing
    }
    elseif($dontcalculate -eq $true){
        $dontcalculate=$false
    }
    else{
        try {
            if($Insights_Ignore_0_values.IsChecked -eq "True"){
                $filtereddatapoints=@()
                $filtereddatapoints=$Insights_data.where{$_.$metricname -ne "0"}
                #write-verbose -message $filtereddatapoints.$metricname
                $yellow=Get-Percentile -Sequence ($filtereddatapoints.$metricname) -Percentile $useforyellow
                $red=Get-Percentile -Sequence ($filtereddatapoints.$metricname) -Percentile $useforred
                $average=[Linq.Enumerable]::average([decimal[]] @($filtereddatapoints.$metricname))
                $median=(get-median $filtereddatapoints.$metricname)
                $Insights_datapoints=$filtereddatapoints.count
            }
            else{
                $yellow=Get-Percentile -Sequence ($Insights_data.$metricname) -Percentile $useforyellow
                $red=Get-Percentile -Sequence ($Insights_data.$metricname) -Percentile $useforred
                $average=[Linq.Enumerable]::average([decimal[]] @($Insights_data.$metricname))
                $median=(get-median $Insights_data.$metricname)
                $Insights_datapoints=$Insights_data.count
            }
            [void]$Insights_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Calculation=$yellowpercentile+"th Percentile";Value=[math]::Round($yellow,2);Datapoints=$Insights_datapoints;Stress_level="Change Yellow to "+[math]::Round($yellow)})
            [void]$Insights_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Calculation=$redpercentile+"th percentile";Value=[math]::Round($red,2);Datapoints=$Insights_datapoints;Stress_level="Change Red to "+[math]::Round($red)})
            if($Insights_Show_Average.IsChecked -eq "True"){
                [void]$Insights_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Calculation="Average";Value=[math]::Round($Average,2);Datapoints=$Insights_datapoints})
            }
            if($Insights_Show_Median.IsChecked -eq "True"){
                [void]$Insights_Datagrid.AddChild([pscustomobject]@{Name=$metricname;Calculation="Median";Value=[math]::Round($Median,2);Datapoints=$Insights_datapoints})
            }
        }
        catch{
            write-verbose -message "Calculation failed this metric probably couldn't be converted to an integer"
        }
    }
}
#endregion

#region prepping the xml
$xaml = $xaml -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[xml]$xml=$xaml
$reader = (New-Object System.Xml.XmlNodeReader $xml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }
#endregion

#Region Image
$image="/9j/4AAQSkZJRgABAQEASABIAAD/4RR3aHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/PiA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA2LjAtYzAwMiA3OS4xNjQ0NjAsIDIwMjAvMDUvMTItMTY6MDQ6MTcgICAgICAgICI+IDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+IDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1M2IChXaW5kb3dzKSIgeG1wOkNyZWF0ZURhdGU9IjIwMjAtMDItMDhUMTc6NDM6NTcrMDE6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDIwLTA5LTA5VDEwOjM5OjM5KzAzOjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDIwLTA5LTA5VDEwOjM5OjM5KzAzOjAwIiBkYzpmb3JtYXQ9ImltYWdlL2pwZWciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDo3NTdkMDkwMS04NmY1LTQwY2QtYThiMC05NjE3OGQyZTExNWYiIHhtcE1NOkRvY3VtZW50SUQ9ImFkb2JlOmRvY2lkOnBob3Rvc2hvcDowNTEyNTZlMy04NTg0LTA0NGYtOGYzYi00YjliY2FhOTA2ZjMiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDoyNjRCMTFBMTMxNEJFQTExQUY2MUYyNzkzNjk1RTk0QiI+IDxwaG90b3Nob3A6RG9jdW1lbnRBbmNlc3RvcnM+IDxyZGY6QmFnPiA8cmRmOmxpPmFkb2JlOmRvY2lkOnBob3Rvc2hvcDo0ZmJkNmI5My04MTEwLWE4NDMtOTJjNS1jZTNjN2UzMDFjNzA8L3JkZjpsaT4gPC9yZGY6QmFnPiA8L3Bob3Rvc2hvcDpEb2N1bWVudEFuY2VzdG9ycz4gPHhtcE1NOkhpc3Rvcnk+IDxyZGY6U2VxPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0iY3JlYXRlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDoyNjRCMTFBMTMxNEJFQTExQUY2MUYyNzkzNjk1RTk0QiIgc3RFdnQ6d2hlbj0iMjAyMC0wMi0wOFQxNzo0Mzo1NyswMTowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjk5MjMxZTg3LTFmZDEtNGM0OS1hYzU4LTRjYmY5M2U5MjU5NyIgc3RFdnQ6d2hlbj0iMjAyMC0wMy0yOVQxNzoxNTo0NSswMzowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIxLjEgKE1hY2ludG9zaCkiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjAyMDkxMjVjLWM0MTktNGU3Zi05ODRmLTJjZTVjOTIwOGYxYyIgc3RFdnQ6d2hlbj0iMjAyMC0wMy0zMFQxMjozMToxOCswMzowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIxLjEgKE1hY2ludG9zaCkiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249ImNvbnZlcnRlZCIgc3RFdnQ6cGFyYW1ldGVycz0iZnJvbSBhcHBsaWNhdGlvbi92bmQuYWRvYmUucGhvdG9zaG9wIHRvIGltYWdlL3BuZyIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0iZGVyaXZlZCIgc3RFdnQ6cGFyYW1ldGVycz0iY29udmVydGVkIGZyb20gYXBwbGljYXRpb24vdm5kLmFkb2JlLnBob3Rvc2hvcCB0byBpbWFnZS9wbmciLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjk5NjZjMzJlLTQxMDEtNGZiNC1iMjQ5LTMzODM5ZDk2YzZjZSIgc3RFdnQ6d2hlbj0iMjAyMC0wMy0zMFQxMjozMToxOCswMzowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIxLjEgKE1hY2ludG9zaCkiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjVlNjAwNTQyLTZiY2EtNGU3MC1hNTJhLTQ1NDU5YzM3ZjM4YyIgc3RFdnQ6d2hlbj0iMjAyMC0wOS0wOVQxMDozOTozOSswMzowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIxLjIgKE1hY2ludG9zaCkiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249ImNvbnZlcnRlZCIgc3RFdnQ6cGFyYW1ldGVycz0iZnJvbSBpbWFnZS9wbmcgdG8gaW1hZ2UvanBlZyIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0iZGVyaXZlZCIgc3RFdnQ6cGFyYW1ldGVycz0iY29udmVydGVkIGZyb20gaW1hZ2UvcG5nIHRvIGltYWdlL2pwZWciLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjc1N2QwOTAxLTg2ZjUtNDBjZC1hOGIwLTk2MTc4ZDJlMTE1ZiIgc3RFdnQ6d2hlbj0iMjAyMC0wOS0wOVQxMDozOTozOSswMzowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIxLjIgKE1hY2ludG9zaCkiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOjVlNjAwNTQyLTZiY2EtNGU3MC1hNTJhLTQ1NDU5YzM3ZjM4YyIgc3RSZWY6ZG9jdW1lbnRJRD0iYWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOmZkYmU5YjdiLWZlNzctZjQ0MS04M2ViLTYyMWE1YzRjYmQzOSIgc3RSZWY6b3JpZ2luYWxEb2N1bWVudElEPSJ4bXAuZGlkOjI2NEIxMUExMzE0QkVBMTFBRjYxRjI3OTM2OTVFOTRCIi8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDw/eHBhY2tldCBlbmQ9InciPz7/7RsIUGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAA8cAVoAAxslRxwCAAACAAAAOEJJTQQlAAAAAAAQzc/6fajHvgkFcHaurwXDTjhCSU0EOgAAAAAA5QAAABAAAAABAAAAAAALcHJpbnRPdXRwdXQAAAAFAAAAAFBzdFNib29sAQAAAABJbnRlZW51bQAAAABJbnRlAAAAAENscm0AAAAPcHJpbnRTaXh0ZWVuQml0Ym9vbAAAAAALcHJpbnRlck5hbWVURVhUAAAAAQAAAAAAD3ByaW50UHJvb2ZTZXR1cE9iamMAAAAMAFAAcgBvAG8AZgAgAFMAZQB0AHUAcAAAAAAACnByb29mU2V0dXAAAAABAAAAAEJsdG5lbnVtAAAADGJ1aWx0aW5Qcm9vZgAAAAlwcm9vZkNNWUsAOEJJTQQ7AAAAAAItAAAAEAAAAAEAAAAAABJwcmludE91dHB1dE9wdGlvbnMAAAAXAAAAAENwdG5ib29sAAAAAABDbGJyYm9vbAAAAAAAUmdzTWJvb2wAAAAAAENybkNib29sAAAAAABDbnRDYm9vbAAAAAAATGJsc2Jvb2wAAAAAAE5ndHZib29sAAAAAABFbWxEYm9vbAAAAAAASW50cmJvb2wAAAAAAEJja2dPYmpjAAAAAQAAAAAAAFJHQkMAAAADAAAAAFJkICBkb3ViQG/gAAAAAAAAAAAAR3JuIGRvdWJAb+AAAAAAAAAAAABCbCAgZG91YkBv4AAAAAAAAAAAAEJyZFRVbnRGI1JsdAAAAAAAAAAAAAAAAEJsZCBVbnRGI1JsdAAAAAAAAAAAAAAAAFJzbHRVbnRGI1B4bEBSAAAAAAAAAAAACnZlY3RvckRhdGFib29sAQAAAABQZ1BzZW51bQAAAABQZ1BzAAAAAFBnUEMAAAAATGVmdFVudEYjUmx0AAAAAAAAAAAAAAAAVG9wIFVudEYjUmx0AAAAAAAAAAAAAAAAU2NsIFVudEYjUHJjQFkAAAAAAAAAAAAQY3JvcFdoZW5QcmludGluZ2Jvb2wAAAAADmNyb3BSZWN0Qm90dG9tbG9uZwAAAAAAAAAMY3JvcFJlY3RMZWZ0bG9uZwAAAAAAAAANY3JvcFJlY3RSaWdodGxvbmcAAAAAAAAAC2Nyb3BSZWN0VG9wbG9uZwAAAAAAOEJJTQPtAAAAAAAQAEgAAAABAAIASAAAAAEAAjhCSU0EJgAAAAAADgAAAAAAAAAAAAA/gAAAOEJJTQQNAAAAAAAEAAAAHjhCSU0EGQAAAAAABAAAAB44QklNA/MAAAAAAAkAAAAAAAAAAAEAOEJJTScQAAAAAAAKAAEAAAAAAAAAAjhCSU0D9QAAAAAASAAvZmYAAQBsZmYABgAAAAAAAQAvZmYAAQChmZoABgAAAAAAAQAyAAAAAQBaAAAABgAAAAAAAQA1AAAAAQAtAAAABgAAAAAAAThCSU0D+AAAAAAAcAAA/////////////////////////////wPoAAAAAP////////////////////////////8D6AAAAAD/////////////////////////////A+gAAAAA/////////////////////////////wPoAAA4QklNBAAAAAAAAAIAADhCSU0EAgAAAAAAAgAAOEJJTQQwAAAAAAABAQA4QklNBC0AAAAAAAYAAQAAAAI4QklNBAgAAAAAABAAAAABAAACQAAAAkAAAAAAOEJJTQQeAAAAAAAEAAAAADhCSU0EGgAAAAADUQAAAAYAAAAAAAAAAAAAAzkAAAJIAAAADgBGAGUAcgBtAGEAbABlAF8ARQB4AHAAZQByAHQAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAkgAAAM5AAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAEAAAAAAABudWxsAAAAAgAAAAZib3VuZHNPYmpjAAAAAQAAAAAAAFJjdDEAAAAEAAAAAFRvcCBsb25nAAAAAAAAAABMZWZ0bG9uZwAAAAAAAAAAQnRvbWxvbmcAAAM5AAAAAFJnaHRsb25nAAACSAAAAAZzbGljZXNWbExzAAAAAU9iamMAAAABAAAAAAAFc2xpY2UAAAASAAAAB3NsaWNlSURsb25nAAAAAAAAAAdncm91cElEbG9uZwAAAAAAAAAGb3JpZ2luZW51bQAAAAxFU2xpY2VPcmlnaW4AAAANYXV0b0dlbmVyYXRlZAAAAABUeXBlZW51bQAAAApFU2xpY2VUeXBlAAAAAEltZyAAAAAGYm91bmRzT2JqYwAAAAEAAAAAAABSY3QxAAAABAAAAABUb3AgbG9uZwAAAAAAAAAATGVmdGxvbmcAAAAAAAAAAEJ0b21sb25nAAADOQAAAABSZ2h0bG9uZwAAAkgAAAADdXJsVEVYVAAAAAEAAAAAAABudWxsVEVYVAAAAAEAAAAAAABNc2dlVEVYVAAAAAEAAAAAAAZhbHRUYWdURVhUAAAAAQAAAAAADmNlbGxUZXh0SXNIVE1MYm9vbAEAAAAIY2VsbFRleHRURVhUAAAAAQAAAAAACWhvcnpBbGlnbmVudW0AAAAPRVNsaWNlSG9yekFsaWduAAAAB2RlZmF1bHQAAAAJdmVydEFsaWduZW51bQAAAA9FU2xpY2VWZXJ0QWxpZ24AAAAHZGVmYXVsdAAAAAtiZ0NvbG9yVHlwZWVudW0AAAARRVNsaWNlQkdDb2xvclR5cGUAAAAATm9uZQAAAAl0b3BPdXRzZXRsb25nAAAAAAAAAApsZWZ0T3V0c2V0bG9uZwAAAAAAAAAMYm90dG9tT3V0c2V0bG9uZwAAAAAAAAALcmlnaHRPdXRzZXRsb25nAAAAAAA4QklNBCgAAAAAAAwAAAACP/AAAAAAAAA4Qk
lNBBQAAAAAAAQAAAADOEJJTQQMAAAAABHPAAAAAQAAAHEAAACgAAABVAAA1IAAABGzABgAAf/Y/+0ADEFkb2JlX0NNAAH/7gAOQWRvYmUAZIAAAAAB/9sAhAAMCAgICQgMCQkMEQsKCxEVDwwMDxUYExMVExMYEQwMDAwMDBEMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMAQ0LCw0ODRAODhAUDg4OFBQODg4OFBEMDAwMDBERDAwMDAwMEQwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCACgAHEDASIAAhEBAxEB/90ABAAI/8QBPwAAAQUBAQEBAQEAAAAAAAAAAwABAgQFBgcICQoLAQABBQEBAQEBAQAAAAAAAAABAAIDBAUGBwgJCgsQAAEEAQMCBAIFBwYIBQMMMwEAAhEDBCESMQVBUWETInGBMgYUkaGxQiMkFVLBYjM0coLRQwclklPw4fFjczUWorKDJkSTVGRFwqN0NhfSVeJl8rOEw9N14/NGJ5SkhbSVxNTk9KW1xdXl9VZmdoaWprbG1ub2N0dXZ3eHl6e3x9fn9xEAAgIBAgQEAwQFBgcHBgU1AQACEQMhMRIEQVFhcSITBTKBkRShsUIjwVLR8DMkYuFygpJDUxVjczTxJQYWorKDByY1wtJEk1SjF2RFVTZ0ZeLys4TD03Xj80aUpIW0lcTU5PSltcXV5fVWZnaGlqa2xtbm9ic3R1dnd4eXp7fH/9oADAMBAAIRAxEAPwD1VJJJJSkkkklKSSSSUpU+q9Vwuk4b8zNfsrb7WtGrnvP0Kqmfn2P/APM3/o1cXH/4xg9uN064n9XbkOa9vi91b/RdH8nZamZZmEJSAsxF02OTwxzcxjxSJEZyo1v9Guf8Z+N6zmfs60Uj/CGxm7X6P6Nu7/z4s6rN6P1G9uPndR6hVU0AVuyfdt/lNyaLNn/XsjF/45VfqljdKuyzkZWVVS7GeGtqtcwC0W12VX6Wvrc6ttb/AMxbmZ0D6otfsptyaaq3MZZkY7nWUVOft9Fl994yWM3b2bv+47HssyPQqtrTs3LxOWWOVyxj5JA+ri4eKE/Rw8eOc/ngw4ecy/dpyxwHL8zx1jFceLJivhlHL7pyTx5IergyYv8A0N1sC93SerU9I+1W5uHk0i2i294ssY4uc1rfWaG+pj3bf0e7+b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"
$icon="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"
$Insights_Image.Source = DecodeBase64Image -ImageBase64 $image
$Monitor_Image.Source = DecodeBase64Image -ImageBase64 $image
$Window.add_Loaded({
    $Window.Icon = DecodeBase64Image -ImageBase64 $icon
})
#EndRegion

#region Monitor_related_actions
$Monitor_Select_Folder.Add_Click({select-Monitor_folderdialog})
$Monitor_SourceType.Add_SelectionChanged({ select-Monitor_inputtype  })
$Monitor_Load_Data.Add_Click({ load-Monitor_data })
$Monitor_export.Add_Click({ export-Monitor_data  })
$Monitor_Copy.Add_Click({ copy-Monitor_data  })
$Monitor_Clear.Add_Click({ Clear-Monitor_data  })
$Monitor_Available_Metrics.Add_SelectionChanged({  get-Monitor_stats })
$Monitor_Folder_Selector.Add_SelectionChanged({ set-Monitor_folderfilter })
$Monitor_Ignore_0_values.Add_Checked({Monitor-clear_selected_metric})
$Monitor_Ignore_0_values.Add_Unchecked({Monitor-clear_selected_metric})
$Monitor_Show_Median.Add_Checked({Monitor-clear_selected_metric})
$Monitor_Show_Median.Add_Unchecked({Monitor-clear_selected_metric})
$Monitor_Show_Average.Add_Checked({Monitor-clear_selected_metric})
$Monitor_Show_Average.Add_Unchecked({Monitor-clear_selected_metric})
#endregion Monitor_related_actions

#region Insights related actions
$Insights_Select_File.Add_Click({Open-Insights_Filedialog})
$Insights_export.Add_Click({ export-Insights_data  })
$Insights_Copy.Add_Click({ copy-Insights_data  })
$Insights_Clear.Add_Click({ Clear-Insights_data  })
$Insights_Available_Metrics.Add_SelectionChanged({ get-Insights_stats })
$Insights_Ignore_0_values.Add_Checked({insights-clear_selected_metric})
$Insights_Ignore_0_values.Add_Unchecked({insights-clear_selected_metric})
$Insights_Show_Median.Add_Checked({insights-clear_selected_metric})
$Insights_Show_Median.Add_Unchecked({insights-clear_selected_metric})
$Insights_Show_Average.Add_Checked({insights-clear_selected_metric})
$Insights_Show_Average.Add_Unchecked({insights-clear_selected_metric})
$Insights_Load_Data.Add_Click({  load-Insights_data  })
#endregion

#region Help actions
$Help_Link1.Add_MouseLeftButtonUp({start-process 'https://support.controlup.com/hc/en-us/articles/360013867718'})
$Help_Link2.Add_MouseLeftButtonUp({start-process 'http://www.youtube.com'})
#endregion


$window.ShowDialog() | out-null
