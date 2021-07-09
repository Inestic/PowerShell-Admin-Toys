<#
  .SYNOPSIS
  WPF gui for Exchange Management Shell Get-MessageTrackingLog commandlet.
  
  .INPUTS
  None.
  
  .OUTPUTS
  None.
  
  .EXAMPLE  
  .\Exchange-2016-MessageTracking-Gui.ps1
  Run the script.
  
  .NOTES
  Designed for Microsoft Exchange 2016.
  Tested on Microsoft Exchange 2016 Version 15.1 (Build 2242.4).
  
  .LINK  
  https://github.com/Inestic/PowerShell-Admin-Toys
  
  VERSION
  v1.0.0
  
  DATE
  08.07.2021
  
  Copyright (c) 2021 Inestic
#>

#Requires -Version 5.1

#region Embedded Resource

#endregion Embedded Resource

#region Variables
#region XAML markup
[xml]$Xaml = '<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
	Name = "Window"
    Title="Microsoft Exchange 2016 Message Tracking"
    MinHeight="300"
    MinWidth="942"
    Width="942"
    IsTabStop="False" FontSize="18"
    FontStretch="Normal" FontStyle="Normal"
    FontWeight="Normal" TextOptions.TextFormattingMode="Ideal"
    TextOptions.TextRenderingMode="Auto" SizeToContent="Height"
    Background="#F1F1F1" Foreground="#262626"
    SnapsToDevicePixels="True">
    <Window.Resources>
        <Storyboard x:Key="HasErrorAnimation"
                    Name="HasErrorAnimation">
            <ColorAnimation Storyboard.TargetName="StatusBarGrid"
                            Storyboard.TargetProperty="(Grid.Background).(SolidColorBrush.Color)"
                            From="#eb3b00"
                            To="#2b579a" 
                            Duration="00:00:02"
							BeginTime="00:00:01" />
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
            <Setter Property="Width" Value="150"/>
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
        <Style TargetType="{x:Type ComboBox}">
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Padding" Value="2" />
            <Setter Property="Width" Value="350" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="IsEditable" Value="False"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid x:Name="templateRoot" SnapsToDevicePixels="True">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition MinWidth="{DynamicResource {x:Static SystemParameters.VerticalScrollBarWidthKey}}" Width="0"/>
                            </Grid.ColumnDefinitions>
                            <Popup x:Name="PART_Popup" Width="350" AllowsTransparency="True" Grid.ColumnSpan="2" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" Margin="1" PopupAnimation="{DynamicResource {x:Static SystemParameters.ComboBoxPopupAnimationKey}}" Placement="Bottom">                                
                                    <Border x:Name="DropDownBorder" BorderBrush="#2b579a" BorderThickness="1" Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}">
                                        <ScrollViewer x:Name="DropDownScrollViewer">
                                            <Grid x:Name="grid" RenderOptions.ClearTypeHint="Enabled">
                                                <Canvas x:Name="canvas" HorizontalAlignment="Left" Height="0" VerticalAlignment="Top" Width="0">
                                                    <Rectangle x:Name="OpaqueRect" Fill="{Binding Background, ElementName=DropDownBorder}" Height="{Binding ActualHeight, ElementName=DropDownBorder}" Width="{Binding ActualWidth, ElementName=DropDownBorder}"/>
                                                </Canvas>
                                                <ItemsPresenter x:Name="ItemsPresenter" KeyboardNavigation.DirectionalNavigation="Contained" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                                            </Grid>
                                        </ScrollViewer>
                                    </Border>                                
                            </Popup>
                            <ToggleButton x:Name="toggleButton" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" Grid.ColumnSpan="2" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                                <ToggleButton.Style>
                                    <Style TargetType="{x:Type ToggleButton}">
                                        <Setter Property="OverridesDefaultStyle" Value="True"/>                                        
                                        <Setter Property="IsTabStop" Value="False"/>
                                        <Setter Property="Focusable" Value="False"/>
                                        <Setter Property="ClickMode" Value="Press"/>                                        
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="{x:Type ToggleButton}">
                                                    <Border x:Name="templateRoot" BorderBrush="#2b579a" BorderThickness="{TemplateBinding BorderThickness}" Background="#FFFFFF" SnapsToDevicePixels="True">                                                        
                                                        <Border x:Name="splitBorder" BorderBrush="Transparent" BorderThickness="1" HorizontalAlignment="Right" Margin="0" SnapsToDevicePixels="True" Width="{DynamicResource {x:Static SystemParameters.VerticalScrollBarWidthKey}}">
                                                            <Path x:Name="Arrow" Data="F1M0,0L2.667,2.66665 5.3334,0 5.3334,-1.78168 2.6667,0.88501 0,-1.78168 0,0z" Fill="#2b579a" HorizontalAlignment="Center" Margin="0" VerticalAlignment="Center"/>
                                                        </Border>
                                                    </Border>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </ToggleButton.Style>
                            </ToggleButton>
                            <Border Padding="2">
                                <ContentPresenter x:Name="contentPresenter" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" Content="{TemplateBinding SelectionBoxItem}" ContentStringFormat="{TemplateBinding SelectionBoxItemStringFormat}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" IsHitTestVisible="False" Margin="{TemplateBinding Padding}" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>                            
                            <Trigger Property="HasItems" Value="False">
                                <Setter Property="Height" TargetName="DropDownBorder" Value="95"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsGrouping" Value="True"/>
                                    <Condition Property="VirtualizingPanel.IsVirtualizingWhenGrouping" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter Property="ScrollViewer.CanContentScroll" Value="False"/>
                            </MultiTrigger>
                            <Trigger Property="CanContentScroll" SourceName="DropDownScrollViewer" Value="False">
                                <Setter Property="Canvas.Top" TargetName="OpaqueRect" Value="{Binding VerticalOffset, ElementName=DropDownScrollViewer}"/>
                                <Setter Property="Canvas.Left" TargetName="OpaqueRect" Value="{Binding HorizontalOffset, ElementName=DropDownScrollViewer}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
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
                             Grid.Column="1"
                             MaxLength="19"
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
                             Grid.Column="1"
                             MaxLength="19"
                             TabIndex="2" />
                </Grid>
            </StackPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="SenderLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Sender" />

                    <TextBox Name="SenderTextBox"
                             Grid.Column="1"
                             TabIndex="3" />
                </Grid>
            </StackPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="RecipientLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Recipient" />

                    <TextBox Name="RecipientTextBox"
                             Grid.Column="1"
                             TabIndex="4" />
                </Grid>
            </StackPanel>
            <StackPanel Style="{StaticResource TextBoxPanel}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="SubjectLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Subject" />

                    <TextBox Name="SubjectTextBox"
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

                    <TextBlock Name="EventIdLabel"
                               Grid.Column="0"
                               Style="{StaticResource LabelTextBlockStyle}"
                               Text="Event ID" />

                    <ComboBox Name="EventIdTextBox"
                              Grid.Column="1"
                              TabIndex="6">
                    </ComboBox>
                </Grid>
            </StackPanel>
        </WrapPanel>
        <Button Name="SearchButton" 
                Margin="105, 0, 0, 0">
            <TextBlock Text="Search"/>
        </Button>
        <Button Name="ExportButton" 
                Margin="305, 0, 0, 0"
                IsEnabled="{Binding Path=IsVisible, ElementName=FoundMailDataGrid}">
            <TextBlock Text="Export"/>
        </Button>
        
        <DataGrid Name="FoundMailDataGrid"/>

        <Grid Name="StatusBarGrid" 
              Style="{StaticResource StatusBarGridStyle}">
            
            <TextBlock Name="StatusBarTextBlock"
                       Style="{StaticResource StatusBarInfoTextStyle}"
					   Text="Ready for mail tracking"/>

            <TextBlock Name="GitHubLink"
                       Style="{StaticResource StatusBarLinkTextStyle}"
                       Text="https://github.com/Inestic"/>            
        </Grid>
    </Grid>
</Window>'
#endregion XAML markup
$MailEventId = [ordered]@{
	"All events" = "ANY"
	"DEFER: message delivery was delayed" = "DEFER"
	"DELIVER:	message was delivered" = "DELIVER"
	"DROP: message was dropped" = "DROP"
	"DSN: delivery status notification created" = "DSN"
	"RECEIVE:	message was received to SMTP" = "RECEIVE"
	"SEND: A message was sent by SMTP" = "SEND"
}
$GitHubPage = "https://github.com/Inestic/PowerShell-Admin-Toys"
$ExchangeSnapin = "Microsoft.Exchange.Management.PowerShell.SnapIn"
#endregion Variables

#region Functions
function Open-Url {
	
	[CmdletBinding()]
		param
		(
			[parameter(Mandatory=$true)]
			[ValidateNotNull()]
			[string]
			$Url
		)
		
	Start-Process $Url
	
}

function Show-Error {
	
	[CmdletBinding()]
		param
		(
			[parameter(Mandatory=$true)]
			[ValidateNotNull()]
			[byte]
			$SequentialNumber
		)
		
	[String[]] $ErrorsDescriptions = "Requires powershell module: Exchange Management Shell"
	$AnimationName = "HasErrorAnimation"
	$StatusBarTextBlock.Text = $ErrorsDescriptions[$SequentialNumber]
	$Animation = $Window.FindResource("$AnimationName")
	$Animation.Begin()
}

function Get-ExchangeSnapin {	

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
	}
}

function Start-MessageTracking {

	[String]$TrackingCommand = [String]::Format("Get-MessageTrackingLog -Start {0} -End {1}", $StartDateTextBox.Text, $EndDateTextBox.Text)
	$SenderTextBox, $RecipientTextBox, $SubjectTextBox
	
	
	#if ([string]::Empty -ne $SenderTextBox.Text) {$TrackingCommand += " -Sender ""{0}"""-f $SenderTextBox.Text}
	
	

}
#endregion Functions

Add-Type -AssemblyName PresentationFramework

$Gui = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $Xaml))
$Xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
	Set-Variable -Name ($_.Name) -Value $Gui.FindName($_.Name)
}

$StartDateTextBox.Text = [DateTime]::Today
$EndDateTextBox.Text = [DateTime]::Today.AddSeconds(-1)
$EventIdTextBox.ItemsSource = $MailEventId.Keys
$EventIdTextBox.SelectedItem = $EventIdTextBox.ItemsSource | Select-Object -First 1
$Window.add_Loaded({Get-ExchangeSnapin})
$GitHubLink.add_MouseLeftButtonDown({Open-Url -Url $GitHubPage})
$SearchButton.add_Click({Start-MessageTracking})
$Gui.ShowDialog() | Out-Null