<#
	.SYNOPSIS
	WPF GUI for collecting and sorting 4625 events from Microsoft Exchange 2016 servers.
	It is now very easy to identify the most frequently blocked users.

	.INPUTS
	None

	.OUTPUTS
	None

	.EXAMPLE Run script
	.\Exchange-2016-MessageTracking-Gui.ps1

	.NOTES
	Designed for Microsoft Exchange 2016
	Tested on Microsoft Exchange 2016 version 15.1 build 2242.4

	.LINK
	https://github.com/Inestic/PowerShell-Admin-Toys

	.VERSION
	v1.0.0

	.DATE
	08.07.2021

	Copyright (c) 2021 Inestic
#>

#Requires -Version 5.1

#region XAML markup
[xml]$Xaml = '<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
	Name = "Window"
    Title="Microsoft Exchange 2016 Bad Password Event Viewer"
    MinHeight="200"
    MinWidth="942"
    Width="942"
    IsTabStop="False"
    FontSize="18"
    FontStretch="Normal"
    FontStyle="Normal"
    FontWeight="Normal"
    TextOptions.TextFormattingMode="Ideal"
    TextOptions.TextRenderingMode="Auto"
    SizeToContent="Height"
    Background="#F1F1F1" Foreground="#262626"
    SnapsToDevicePixels="True">
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
        <Style TargetType="{x:Type Button}">
            <Setter Property="Grid.Row" Value="1"/>
            <Setter Property="Height" Value="30"/>
            <Setter Property="Width" Value="120"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Left"/>
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
            <Style.Triggers>
                <DataTrigger Binding="{Binding ElementName=EventDataGrid, Path=Items.Count}" Value="0">
                    <Setter Property="Visibility" Value="Collapsed"/>
                </DataTrigger>
            </Style.Triggers>
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
            <Setter Property="Width" Value="300" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Grid.Column" Value="1"/>
            <Setter Property="MaxLength" Value="19"/>
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
                    <TextBlock Name="StartLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Start Date" />
                    <TextBox Name="StartDateTextBox"                             
                             TabIndex="1" />
                </Grid>
            </StackPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="EndLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="End Date" />
                    <TextBox Name="EndDateTextBox"                             
                             TabIndex="2" />
                </Grid>
            </StackPanel>
        </WrapPanel>
        <Button Name="SearchButton" 
                Margin="105, 0, 0, 0"
                Content="Search"/>

        <Button Name="ActionButton" 
                Margin="280, 0, 0, 0"
                Content="Change view"
                Visibility="{Binding ElementName=EventDataGrid, Path=Visibility}"/>

        <DataGrid Name="EventDataGrid"/>
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
$ExchangeSnapin = "Microsoft.Exchange.Management.PowerShell.SnapIn"
$FailureEvents = New-Object System.Collections.ArrayList($null)
$FailureUsers = New-Object System.Collections.ArrayList($null)
$SecurityLog = "Security"
$EventID = "4625"
$EventShema = "Microsoft-Windows-Security-Auditing"
$LogonTypes = @{
	[uint32]2 = "Interactive"
	[uint32]3 = "Network"
	[uint32]4 = "Batch"
	[uint32]5 = "Service"
	[uint32]7 = "Unlock"
	[uint32]8 = "NetworkCleartext"
	[uint32]9 = "NewCredentials"
	[uint32]10 = "RemoteInteractive"
	[uint32]11 = "CachedInteractive"
 }
#endregion Variables

#region Functions
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

	[String[]] $ErrorsDescriptions = "Requires powershell module: Exchange Management Shell"
	$StatusBarTextBlock.Text = $ErrorsDescriptions[$SequentialNumber]
}
function Get-ExchangeSnapin
{
	try
	{
		if ((Get-PSSnapin).Name -notcontains $ExchangeSnapin)
		{
			Add-PSSnapin -Name $ExchangeSnapin -ErrorAction Stop
		}
	}
	catch
	{
		$SearchButton.IsEnabled = $false
		Show-Error -SequentialNumber 0
		Show-Animation
	}
}

function Search-FailureEvents
{	
	$FilterData = @{LogName=$SecurityLog;ID=$EventID;StartTime=$StartDateTextBox.Text;EndTime=$EndDateTextBox.Text;ProviderName=$EventShema}
	Get-ExchangeServer | ForEach-Object -Process {
		$StatusBarTextBlock.Text = "Search started, please wait"
		$ExchangeServer = $_		
		Get-WinEvent -ComputerName $ExchangeServer.Name -FilterHashtable $FilterData | ForEach-Object -Process {
			$Property = [ordered]@{}
			$Property.TimeCreated     = $_.TimeCreated
			$Property.AccountName     = $_.Properties[5].Value
			$Property.AccountDomain   = $_.Properties[6].Value
			$Property.WorkstationName = $_.Properties[13].Value
			$Property.NetworkAddress  = $_.Properties[19].Value
			$Property.Port            = $_.Properties[20].Value
			$Property.LogonType       = $LogonTypes[$_.Properties[10].Value]
			$Property.ProcessName     = $_.Properties[18].Value
			$Property.Server          = $_.MachineName
			[Void]$FailureEvents.Add((New-Object -TypeName PSObject -Property $Property))			
		}
	}
	
	$FailureEvents | Group-Object -Property AccountName | Sort-Object -Property Count -Descending | ForEach-Object -Process {
		$Property = [ordered]@{}
		$Property.BadLogonCount = $_.Count
		$Property.AccountName   = $_.Name
		[Void]$FailureUsers.Add((New-Object -TypeName PSObject -Property $Property))
	}		
	
	$EventDataGrid.ItemsSource = $FailureUsers
	$StatusBarTextBlock.Text = "Find: {0}"-f $FailureUsers.Count
	Show-Animation
}

function Change-DataView
{
	if ($EventDataGrid.ItemsSource.Count -eq $FailureEvents.Count)
	{
		$EventDataGrid.ItemsSource = $FailureUsers
		$StatusBarTextBlock.Text = "Find: {0}"-f $FailureUsers.Count
	}
	else
	{
		$EventDataGrid.ItemsSource = $FailureEvents | Sort-Object -Property AccountName
		$StatusBarTextBlock.Text = "Find: {0}"-f $FailureEvents.Count
	}
}
#endregion Functions

Add-Type -AssemblyName PresentationFramework

$Gui = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $Xaml))
$Xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
	Set-Variable -Name ($_.Name) -Value $Gui.FindName($_.Name)
}
$Today = [DateTime]::Today
$StartDateTextBox.Text       = "{0} {1}"-f $Today.ToShortDateString(), $Today.ToLongTimeString()
$EndDateTextBox.Text         = "{0} {1}"-f $Today.ToShortDateString(), $Today.AddDays(1).AddSeconds(-1).ToLongTimeString()
$Window.Add_Loaded({Get-ExchangeSnapin})
$GitHubLink.Add_MouseLeftButtonDown({Open-Url -Url $GitHubLink.Text})
$SearchButton.Add_Click({Search-FailureEvents})
$ActionButton.Add_Click({Change-DataView})
$Gui.ShowDialog() | Out-Null
