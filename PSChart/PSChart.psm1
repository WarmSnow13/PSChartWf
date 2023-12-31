[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')
# Add-Type -AssemblyName System.Drawing

<#
.Synopsis
    Create new oblect type of System.Windows.Forms.DataVisualization.Charting.Series

.DESCRIPTION
    Defines functions for wokring with  Microsoft Chart Control for .NET 3.5 Framework
    Note: Requires NET Framework 3.5 and Microsoft Chart Controls for Microsoft .NET Framework

.EXAMPLE

    Import-Module PSChart
    Import-IseSnippet -Module PSChart

    $dataSource = Get-Process |select Name,VirtualMemorySize,WorkingSet,PeakWorkingSet | Select-Object -First 6
    $Series1 = New-ChartSeries -XValueMember Name -YValueMembers VirtualMemorySize -ChartType Column
    $Series2 = New-ChartSeries -XValueMember Name -YValueMembers WorkingSet -ChartType Column
    $Series3 = New-ChartSeries -XValueMember Name -YValueMembers PeakWorkingSet -ChartType Spline
    New-Chart -datasource $dataSource -series $Series1,$Series2,$Series3 -SaveAs C:\temp\result.png -Enable3D

.EXAMPLE
    $dataSource = Get-Process |select Name,VirtualMemorySize,WorkingSet,PeakWorkingSet  | Select-Object -First 6
    New-Chart -ChartType spline -Properties Name,VirtualMemorySize -DataSource $dataSource

.EXAMPLE
    $dataSource = ls C:\temp | where { $_.PSIsContainer -EQ $false} | Select-Object -First 4
    $dataSource | New-Chart -ValueMembers name,LastAccessTime -ChartType bar

.EXAMPLE
    $dataSource = ls C:\temp | where { $_.PSIsContainer -EQ $false} | Select-Object -First 3
    $Series1 = New-ChartSeries -XValueMember name -YValueMembers LastAccessTime -ChartType column -XValueType DateTime
    New-Chart -datasource $dataSource -series $Series1 -SaveAs C:\temp\foo -Silent -Clipboard

#>
Function New-ChartSeries {

    #region Parameters
    param(
        [parameter(Mandatory = $true)]
        [string]
        $XValueMember,

        [parameter(Mandatory = $true)]
        [string]
        $YValueMembers,

        #region Optional Parameters

        [parameter(Mandatory = $false)]
        [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
        $ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column,

        [parameter(Mandatory = $false)]
        [System.Windows.Forms.DataVisualization.Charting.ChartValueType]
        $XValueType,
    
        [parameter(Mandatory = $false)]
        [System.Windows.Forms.DataVisualization.Charting.ChartValueType]
        $YValueType

        #endregion

    )
    #endregion

    [void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')
    $Series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    # [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $Series.ChartType = $ChartType
    $Series.XValueMember = $XValueMember;
    $Series.YValueMembers = $YValueMembers;

    $Series.LegendToolTip = "$YValueMembers"
    $Series.IsVisibleInLegend = $true
    $Series.LegendText = "$YValueMembers"

    $Series.IsValueShownAsLabel = $true
    #$Series.LabelFormat = "N2"
    $Series.LabelBackColor = [System.Drawing.Color]::WhiteSmoke
    $Series.Font = New-Object system.drawing.font('calibri', 10, [system.drawing.fontstyle]::Regular)

    $Series.LabelToolTip = "$YValueMembers"

    if ($XValueType) { 
        $Series.XValueType = $XValueType;
    }
    if ($YValueType) { 
        $Series.YValueType = $YValueType;
    }

    #$Series.Points.DataBind( $e1, 'x', 'y', "");
    return $Series
}


<#
.Synopsis
   Create the chart object

.DESCRIPTION
    Defines functions for wokring with Microsoft Chart Control for .NET 3.5 Framework
    Note: Requires NET Framework 3.5 and Microsoft Chart Controls for Microsoft .NET Framework

.EXAMPLE
    Get-Process | select -First 3 | New-Chart -ValueMembers Name,VirtualMemorySize -ChartType Pie -Enable3D  

.EXAMPLE
    Access to Online Documentation
    Get-Help New-Chart -online

.EXAMPLE
    Import-Module PSChart
    $dataSource = Get-Process |select Name,VirtualMemorySize,WorkingSet,PeakWorkingSet  | Select-Object -First 6
    $Series1 = New-ChartSeries -XValueMember Name -YValueMembers VirtualMemorySize -ChartType Column
    $Series2 = New-ChartSeries -XValueMember Name -YValueMembers WorkingSet -ChartType Column
    $Series3 = New-ChartSeries -XValueMember Name -YValueMembers PeakWorkingSet -ChartType Spline
    New-Chart -datasource $dataSource -series $Series1,$Series2,$Series3 -SaveAs C:\temp\image.png -Enable3D

.EXAMPLE
    $dataSource = Get-Process |select Name,VirtualMemorySize,WorkingSet,PeakWorkingSet  | Select-Object -First 6
    New-Chart -ChartType spline -Properties Name,VirtualMemorySize -DataSource $dataSource

#>
function New-Chart {
    [CmdletBinding(
        DefaultParametersetName = 'Single'
    )]

    #region Parameters
    param(
        [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
        [ValidateCount(2, 2)]
        [Alias('Properties', 'Members')]
        [string[]]$ValueMembers,

        [Parameter(ParameterSetName = 'Multiple', Mandatory = $true)]
        [System.Windows.Forms.DataVisualization.Charting.Series[]]
        $Series,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [Parameter(ValueFromPipeline = $true)]
        #[Parameter(ValueFromPipeline=$true,Position=0)]
        $DataSource,

        #region Optional Parameters
        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [switch]
        $Enable3D,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [string]
        $SaveAs,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [string]
        $ChartTitle,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [string]
        $YTitle, $XTitle,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [switch]
        $Clipboard, 

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'Multiple')]
        [parameter(Mandatory = $false)]
        [switch]
        $Silent
        <#
        # moved to DynamicParam
        [Parameter(ParameterSetName="Single")]
        [Parameter(Mandatory=$false,Position=2)]
        [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
        $ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::bar,
        #>

    )
    #endregion


    DynamicParam {
        $Attributes = New-Object 'Management.Automation.ParameterAttribute'
        $Attributes.ParameterSetName = 'Single'
        $Attributes.Mandatory = $false
        $Attributes.Position = 2
        
        $AttributesCollection = New-Object 'Collections.ObjectModel.Collection[Attribute]'
        $AttributesCollection.Add($Attributes)
        
        $ChartType = New-Object -TypeName 'Management.Automation.RuntimeDefinedParameter' -ArgumentList @('ChartType', [System.Windows.Forms.DataVisualization.Charting.SeriesChartType], $AttributesCollection)

        $ParamDictionary = New-Object 'Management.Automation.RuntimeDefinedParameterDictionary'
        $ParamDictionary.Add('ChartType', $ChartType)
        $ParamDictionary
    }

    Begin {
        $data = @()
    }
    Process {
        $data += $datasource
    }
    End {

        #region chartarea

        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        #$ChartArea.BackColor = [System.Drawing.Color]::Transparent
        $ChartArea.Area3DStyle.Enable3D = $Enable3D
        $ChartArea.Area3DStyle.Inclination = 45
        $ChartArea.Area3DStyle.LightStyle = [System.Windows.Forms.DataVisualization.Charting.LightStyle]::Realistic

        $font = New-Object system.drawing.font('calibri', 10, [system.drawing.fontstyle]::Regular)

        $ChartArea.AxisX.Titlefont = $font
        $ChartArea.AxisY.Titlefont = $font

        $ChartArea.AxisX.LabelAutoFitStyle = [System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles]::StaggeredLabels
        $ChartArea.AxisX.LabelStyle.Font = $font

        $ChartArea.AxisY.LabelStyle.Format = 'N0'   #ref: http://msdn.microsoft.com/en-us/library/dwhawy9k.aspx
        $ChartArea.AxisY.LabelStyle.Font = $font

        #$ChartArea1 = Import-Clixml
        #$ChartArea.BackColor = ([System.Drawing.Color]::DarkOrange)
        #$ChartArea.AxisX.LineColor = ([System.Drawing.Color]::Green)

        if ($XTitle) { $chartarea.AxisX.Title = $XTitle; }
        if ($YTitle) { $chartarea.AxisY.Title = $YTitle; }
  
        $chartarea.AxisX.Interval = 1             
        #$ChartArea.AxisY.IsLogarithmic = $true

        #endregion

        #region Legends
        $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
        $Legend.name = 'Legend1'
        #endregion

        $Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $Chart.Width = 800; 
        $Chart.Height = 600 #[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height - 40 # 500
        $Chart.Left = 0; 
        $Chart.Top = 20;
        $Chart.BackColor = [System.Drawing.Color]::White # [System.Drawing.Color]::Transparent

        $Chart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
        [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

        $Chart.Dock = [System.Windows.forms.DockStyle]::None
        $Chart.AutoSize = $false;

        #region ChartTitle   
        if ($ChartTitle) {                                                      
            [void]$chart.Titles.Add("$ChartTitle");
            $chart.Titles[0].Font = 'Arial,13pt';                     
            $chart.Titles[0].Alignment = 'topLeft';
        }                      
        #endregion

        $Chart.ChartAreas.Add($ChartArea)
        $Chart.Legends.Add($Legend)

        #Chart_Series: add series to chart 
        switch ($PsCmdlet.ParameterSetName) {
            'Single' { 
                Write-Debug 'Single Chart'; 
                [string]$XValueMember = $ValueMembers[0];
                [string]$YValueMember = $ValueMembers[1];

                #($ChartType.IsSet -eq $false)

                $Series1 = if ($ChartType.value) {
                    New-ChartSeries -XValueMember $XValueMember -YValueMembers $YValueMember -ChartType $ChartType.value
                }
                else {
                    New-ChartSeries -XValueMember $XValueMember -YValueMembers $YValueMember
                }

                #$Series1.Points.Clear();
                $data | ForEach-Object {
                    $Series1.Points.addxy( $_.($Series1.XValueMember), $_.($Series1.YValueMembers) )
                }
    
                [void]$Chart.Series.Add($Series1)
           
                break;
            }
            'Multiple' { 
                Write-Debug 'Multiple Charts'; 
                foreach ($item in $series) {

                    $data | ForEach-Object {
                        [void]$item.Points.addxy( $_.($item.XValueMember), $_.($item.YValueMembers) )
                    }

                    [void]$Chart.Series.Add($item)
                }
                break;
            }
        }

        $datasource
        #$Chart.DataSource = $data;
        #$Chart.DataBind();

        #region processing switch 'SaveAs' 
        if ($SaveAs) {
            if ($SaveAs.ToLower().EndsWith('.emf')) {
                $Chart.SaveImage($SaveAs, [System.Drawing.Imaging.ImageFormat]::Emf)
            }
            if (!($SaveAs.ToLower().EndsWith('.emf')) -or ( [System.IO.Path]::GetExtension($SaveAs) -eq '') ) {
                [string]$newName = [System.IO.Path]::Combine(
                    [System.IO.Path]::GetDirectoryName($SaveAs),
                    [System.IO.Path]::GetFileNameWithoutExtension($SaveAs) + '.png'
                );

                $Chart.SaveImage($newName, [System.Drawing.Imaging.ImageFormat]::Png)
            }
        }
        #endregion 

        #region processing switch 'Clipboard' 
        if ($Clipboard) {
            $newms = New-Object System.IO.MemoryStream
            $Chart.SaveImage($newms, [System.Drawing.Imaging.ImageFormat]::png)
            [System.Drawing.Image]$image = [System.Drawing.Image]::FromStream($newms)
            [Windows.Forms.Clipboard]::SetImage($image)
        }
        #endregion

        #region processing switch 'Silent' 
        if ($Silent -eq $false) {
            [System.Windows.Forms.Application]::EnableVisualStyles();

            $Form = New-Object Windows.Forms.Form -Property @{
                BackColor     = [System.Drawing.Color]::FromArgb(255, 255, 255, 255);
                StartPosition = 1;
                Width         = 800 + 250;
                Height        = 600 + 100 #[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height;
                Text          = 'PSChart';
            }

            #region SaveButton
            $SaveButton = New-Object Windows.Forms.Button
            $SaveButton.FlatStyle = 0
            $SaveButton.Text = 'Save Chart As'
            $SaveButton.Top = 1
            $SaveButton.Left = 1
            $SaveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    
            $SaveButton.add_click( {

                    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $SaveFileDialog.initialDirectory = $Env:USERPROFILE + '\Desktop\'
                    $SaveFileDialog.filter = 'All files (*.*)| *.*'
                    $saveFileDialog.FileName = Get-Date -Format 'yyyymmdd_hhmmss'
                    $SaveFileDialog.DefaultExt = 'png'

                    [void]$SaveFileDialog.ShowDialog()
                    Write-Debug $SaveFileDialog.filename
                    if ($SaveFileDialog.filename) {
                        $Chart.SaveImage($SaveFileDialog.filename, [System.Drawing.Imaging.ImageFormat]::png)
                    }
                })
            [void]$Form.controls.add($SaveButton)
            #endregion

            #region SaveButton .... 
            $SaveButton = New-Object Windows.Forms.Button
            $SaveButton.FlatStyle = 0
            $SaveButton.Text = 'Save Style'
            $SaveButton.Top = 1
            $SaveButton.Left = 300
            $SaveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    
            $SaveButton.add_click( {
                    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $SaveFileDialog.initialDirectory = $Env:USERPROFILE + '\Desktop\'
                    $SaveFileDialog.filter = 'All files (*.*)| *.xml'
                    $saveFileDialog.FileName = Get-Date -Format 'yyyymmdd_hhmmss'
                    $SaveFileDialog.DefaultExt = 'xml'

                    [void]$SaveFileDialog.ShowDialog()
                    Write-Debug $SaveFileDialog.filename
                    if ($SaveFileDialog.filename) {
                        $ChartArea | Export-Clixml -Depth 5 -Path ($SaveFileDialog.filename)
                    }
                })
 
            [void]$Form.controls.add($SaveButton)
            #endregion

            #region PropertyGrid
            $propertyGrid1 = New-Object System.Windows.Forms.PropertyGrid
            $propertyGrid1.DataBindings.DefaultDataSourceUpdateMode = 0
            $propertyGrid1.Dock = [System.Windows.forms.DockStyle]::Right #4 # [System.Windows.forms.DockStyle]::Fill 
    
            $propertyGrid1.Font = New-Object System.Drawing.Font('Segoe UI Light', 8.25, 0, 3, 0)
            $propertyGrid1.Location = [System.Drawing.Point] @{X = 471; Y = 0 }
            $propertyGrid1.SelectedObject = $Chart
            $propertyGrid1.Size = [System.Drawing.Size] @{Height = 500; Width = 250; }
            [void]$Form.Controls.Add($propertyGrid1)
            #endregion 

            #region SeriesFormatSelector
            $comboBox1 = New-Object System.Windows.Forms.ComboBox -Property @{ DropDownStyle = 2; FlatStyle = 0; }
            $comboBox1.Location = New-Object System.Drawing.Point -Property @{ X = 100; Y = 1; }
            $comboBox1.Size = New-Object System.Drawing.Size -Property @{ Height = 21; Width = 80; }

            for ($i = 0; $i -lt $chart.Series.Count; $i++) { [void]$comboBox1.Items.Add($chart.Series[$i].Name) }

            $comboBox1.SelectedIndex = 0
            $comboBox1.add_SelectedIndexChanged({
                $propertyGrid1.SelectedObject = $Chart.Series[$comboBox1.SelectedIndex] 
            })
            [void]$Form.Controls.Add($comboBox1)
            #endregion

            #region FormatSelector
            $comboBox2 = New-Object System.Windows.Forms.ComboBox
            $comboBox2.DropDownStyle = 2
            $comboBox2.FlatStyle = 0

            [void]$comboBox2.Items.Add('Chart')
            [void]$comboBox2.Items.Add('Chart Area')
            $comboBox2.SelectedIndex = 0

            $comboBox2.Location = New-Object System.Drawing.Point -Property @{ X = 180; Y = 1; }
            $comboBox2.Size = New-Object System.Drawing.Size -Property @{ Height = 21; Width = 80; }

            $comboBox2.add_SelectedIndexChanged( {
                    if ($comboBox2.SelectedIndex -eq 0) { $propertyGrid1.SelectedObject = $Chart }
                    if ($comboBox2.SelectedIndex -eq 1) { $propertyGrid1.SelectedObject = $ChartArea }
                }
            )
            [void]$form.Controls.Add($comboBox2)
            #endregion

            if (-not $Silent) {
                [void]$Form.controls.add($Chart)
                [void]$Form.Add_Shown({ $Form.Activate() })
                [void]$Form.ShowDialog()
            }

            #$Chart.Dispose()
        }
        #endregion

    }
}

Export-ModuleMember -function New-Chart, New-ChartSeries
