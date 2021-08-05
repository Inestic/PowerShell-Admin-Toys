<#
	.SYNOPSIS
	WPF GUI for VMware PowerCLI Get-VIPermission commandlet.	

	.INPUTS
	None

	.OUTPUTS
	None

	.EXAMPLE Run script
	.\vSphere-Virtual-Machines-Permissions-Viewer.ps1

	.NOTES
	Designed for VMware vSphere.
	Tested on VMware vSphere version 6.7.0 build 14368027

	.LINK
	https://github.com/Inestic/PowerShell-Admin-Toys

	.VERSION
	v1.0.0

	.DATE
	05.08.2021

	Copyright (c) 2021 Inestic
#>

#Requires -Version 5.1

#region XAML markup
[xml]$Xaml = '<Window
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"    
	Name="Window"
    Title="vSphere 6.7.0 Virtual Machines Permissions Viewer"   
    IsTabStop="False"
    FontFamily="Cambria"
    FontSize="18"
    FontStretch="Normal" 
    FontStyle="Normal"
    FontWeight="Normal" 
    TextOptions.TextFormattingMode="Ideal"
    TextOptions.TextRenderingMode="Auto" 
    SizeToContent="Height"
    Background="#F1F1F1" 
    Foreground="#262626"
    SnapsToDevicePixels="True"
    Width="596"
    Height="266"
    MinWidth="596"
    MinHeight="266">
    <Window.Resources>
        <Storyboard x:Key="FlashAnimation"
                    Name="FlashAnimation">
            <ColorAnimation Storyboard.TargetName="StatusBarGrid"
                            Storyboard.TargetProperty="(Grid.Background).(SolidColorBrush.Color)"
                            From="#eb3b00"
                            To="#2b579a" 
                            Duration="00:00:0.5"/>
        </Storyboard>
        <Style TargetType="{x:Type Grid}" x:Key="StatusBarGridStyle">
            <Setter Property="Grid.Row" Value="3" />
            <Setter Property="Background">
                <Setter.Value>
                    <SolidColorBrush Color="#2b579a"/>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}" x:Key="StatusBarInfoTextStyle">
            <Setter Property="Margin" Value="10, 0, 0, 0" />
            <Setter Property="Focusable" Value="False" />
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Left"/>
            <Setter Property="Foreground" Value="#e9f3fa" />            
            <Style.Triggers>
                <DataTrigger Binding="{Binding ElementName=PasswordTextBox, Path=Password.Length}" Value="0">
                    <Setter Property="TextBlock.Text" Value="Enter vCenter password"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding ElementName=LoginTextBox, Path=Text.Length}" Value="0">
                    <Setter Property="TextBlock.Text" Value="Enter vCenter login"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding ElementName=vCenterTextBox, Path=Text.Length}" Value="0">
                    <Setter Property="TextBlock.Text" Value="Enter vCenter name"/>
                </DataTrigger>                
            </Style.Triggers>            
        </Style>
        <Style TargetType="{x:Type TextBlock}" x:Key="StatusBarLinkTextStyle">
            <Setter Property="Margin" Value="0, 0, 10, 0" />
            <Setter Property="Focusable" Value="False" />
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Right"/>
            <Setter Property="Foreground" Value="#e9f3fa" />
            <Setter Property="FontSize" Value="14"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="TextDecorations" Value="Underline"/>
                    <Setter Property="ForceCursor" Value="True"/>
                    <Setter Property="Cursor" Value="Hand"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type TextBlock}" x:Key="LabelTextBlockStyle">
            <Setter Property="Margin" Value="10, 5, 10, 5" />
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="Focusable" Value="False" />
        </Style>
        <Style TargetType="{x:Type Button}" x:Key="SimpleButton">
            <Setter Property="Grid.Row" Value="1"/>
            <Setter Property="Height" Value="30"/>
            <Setter Property="Width" Value="150"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Left"/>
            <Setter Property="Margin" Value="305, 0, 0, 0"/>
            <Setter Property="Content" Value="Export"/>            
        </Style>
        <Style TargetType="{x:Type Button}" BasedOn="{StaticResource SimpleButton}" x:Key="BindingButton">
            <Setter Property="IsEnabled" Value="False"/>
            <Setter Property="Margin" Value="105, 0, 0, 0"/>
            <Setter Property="Content" Value="Connect"/>            
        </Style>
        <Style TargetType="{x:Type WrapPanel}">
            <Setter Property="Grid.Row" Value="0" />
            <Setter Property="Orientation" Value="Horizontal" />
        </Style>
        <Style TargetType="{x:Type DataGrid}">
            <Setter Property="Grid.Row" Value="2" />
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="AutoGenerateColumns" Value="True"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="GridLinesVisibility" Value="All"/>
            <Setter Property="CanUserReorderColumns" Value="False"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="ColumnWidth" Value="*"/>
            <Setter Property="Visibility" Value="Collapsed"/>
        </Style>
        <Style TargetType="{x:Type StackPanel}" x:Key="TextBoxPanel">
            <Setter Property="Orientation" Value="Horizontal" />
            <Setter Property="Margin" Value="5" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Padding" Value="2" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#2b579a" />
            <Setter Property="Width" Value="350" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Background" Value="#FFFFFF"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Padding" Value="2" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#2b579a" />
            <Setter Property="Width" Value="350" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Background" Value="#FFFFFF"/>            
        </Style>
    </Window.Resources>    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="50" />
            <RowDefinition Height="*" />
            <RowDefinition Height="40" />
        </Grid.RowDefinitions>
        <WrapPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="vCenterLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="vCenter" />

                    <TextBox Name="vCenterTextBox"
                             Grid.Column="1"                             
                             TabIndex="1" />
                </Grid>
            </StackPanel>            
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="LoginLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Login" />

                    <TextBox Name="LoginTextBox"
                             Grid.Column="1"
                             TabIndex="5" />
                </Grid>
            </StackPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>

                    <TextBlock Name="PasswordLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Password" />

                    <PasswordBox Name="PasswordTextBox"
                                 Grid.Column="1"
                                 TabIndex="6"/>
                </Grid>
            </StackPanel>
        </WrapPanel>
        <Button Name="ConnectButton"                
                Style="{StaticResource BindingButton}"/>
        <Button Name="ExportButton"
                Style="{StaticResource SimpleButton}"
                IsEnabled="{Binding ElementName=VmsDataGrid, Path=IsVisible}"/>
        <DataGrid Name="VmsDataGrid"/>
        <Grid Name="StatusBarGrid" 
              Style="{StaticResource StatusBarGridStyle}">

            <TextBlock Name="StatusBarTextBlock"
                       Style="{StaticResource StatusBarInfoTextStyle}"/>

            <TextBlock Name="GitHubLink"
                       Style="{StaticResource StatusBarLinkTextStyle}"
                       Text="https://github.com/Inestic"/>
        </Grid>
    </Grid>
</Window>'
#endregion XAML markup

#region Variables
$GitHubPage = "https://github.com/Inestic/PowerShell-Admin-Toys"
$PowerCliName = "VMware.VimAutomation.Core"
#endregion Variables

#region Functions
function Export-VmData
{
	$SaveFileDialog = New-Object Microsoft.Win32.SaveFileDialog($null)
	$SaveFileDialog.FileName = "VMPermissions_{0}"-f [DateTime]::Now.ToShortDateString()
	$SaveFileDialog.DefaultExt = ".csv"
	$SaveFileDialog.Filter = "Comma separated values (.csv)|*.csv"

	if ($SaveFileDialog.ShowDialog())
	{
		$VmsDataGrid.Items | ConvertTo-Csv -Delimiter "*" -NoTypeInformation | Set-Content -Path $SaveFileDialog.FileName -Encoding Default -Force
	}	
}

function Get-VmData
{
	$Password = ConvertTo-SecureString $PasswordTextBox.Password -AsPlainText -Force
	$Credential = New-Object System.Management.Automation.PSCredential ($LoginTextBox.Text, $Password)	
	
	try
	{
		Connect-VIServer -Server $vCenterTextBox.Text -Credential $Credential -Force -ErrorAction Stop
		$VMData = New-Object System.Collections.ArrayList($null)
		Get-VIPermission | Where-Object {$_.EntityId.Contains("VirtualMachine")} | ForEach-Object -Process {
			$Property = [ordered]@{}			
			$Property.VirtualMachine = $_.Entity.Name
			$Property.UserName       = $_.Principal.ToLower()
			$Property.AccessRights   = $_.Role
			[Void]$VMData.Add((New-Object -TypeName PSObject -Property $Property))
		}
		
		Disconnect-VIServer -Confirm:$false
		
		if ($VMData.Count -gt 0)
		{
			$VmsDataGrid.ItemsSource = $VMData | Sort-Object -Property UserName
			$VmsDataGrid.Visibility = "Visible"
			$StatusBarTextBlock.Text = "Found: {0}"-f $VMData.Count
			Show-Animation
		}
	}
	
	catch
	{
		Show-Error -SequentialNumber 1
		Show-Animation
	}
}

function Verify-Password
{
	if ($PasswordTextBox.Password.Length -gt 0)
	{
		$ConnectButton.IsEnabled = $true
		$StatusBarTextBlock.Text = "Ready connect to vCenter"
	}
	
	else
	{
		$ConnectButton.IsEnabled = $false
		$StatusBarTextBlock.Text = "Enter vCenter password"
	}
}

function Open-Url
{
	[CmdletBinding()]
		param
		(
			[parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[string]
			$Url
		)

	Start-Process -FilePath $Url
}

function Show-Animation
{
	$FlashAnimation.Begin()
}

function Show-Error
{
	[CmdletBinding()]
		param
		(
			[parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[byte]
			$SequentialNumber
		)

	[String[]] $ErrorsDescriptions = "Requires powershell module: VMware PowerCLI", "Oops! Something went wrong..."
	$StatusBarTextBlock.Text = $ErrorsDescriptions[$SequentialNumber]
	
}

function Get-PowerCliModule
{
	try
	{
		if ((Get-Module).Name -notcontains $PowerCliName)
		{
			Import-Module -Name $PowerCliName -ErrorAction Stop
		}
		
		$vCenterTextBox.Focus() 
	}
	catch
	{
		Show-Error -SequentialNumber 0
		Show-Animation
	}
}
#endregion Functions

Add-Type -AssemblyName PresentationFramework

$Gui = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $Xaml))
$Xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
	Set-Variable -Name ($_.Name) -Value $Gui.FindName($_.Name)
}

$Window.Add_Loaded({Get-PowerCliModule})
$GitHubLink.Add_MouseLeftButtonDown({Open-Url -Url $GitHubPage})
$PasswordTextBox.add_PasswordChanged({Verify-Password})
$ConnectButton.add_Click({Get-VmData})
$ExportButton.Add_Click({Export-VmData})
$Gui.ShowDialog() | Out-Null