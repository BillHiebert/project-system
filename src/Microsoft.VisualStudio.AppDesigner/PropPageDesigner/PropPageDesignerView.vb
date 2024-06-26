﻿' Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

Imports Microsoft.VisualStudio.Editors.AppDesCommon
Imports Microsoft.VisualStudio.Editors.AppDesCommon.Utils
Imports Microsoft.VisualStudio.Editors.AppDesDesignerFramework
Imports Microsoft.VisualStudio.Editors.AppDesInterop
Imports Microsoft.VisualStudio.Editors.ApplicationDesigner
Imports Microsoft.VisualStudio.Editors.PropertyPages
Imports Microsoft.VisualStudio.ManagedInterfaces.ProjectDesigner
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Utilities

Imports OleInterop = Microsoft.VisualStudio.OLE.Interop
Imports VSITEMID = Microsoft.VisualStudio.Editors.VSITEMIDAPPDES

Namespace Microsoft.VisualStudio.Editors.PropPageDesigner

    ''' <summary>
    ''' This is the UI for the PropertyPageDesigner
    ''' The view implements IVsProjectDesignerPageSite to allow the property page to 
    ''' notify us of property changes.  The page then sends private change notifications
    ''' which let us bubble the notification into the standard component changed mechanism.
    ''' This will cause the normal undo mechanism to be invoked.
    ''' </summary>
    Public NotInheritable Class PropPageDesignerView
        Inherits UserControl
        Implements IVsProjectDesignerPageSite
        Implements IVsWindowPaneCommit
        Implements IVsEditWindowNotify
        Implements IServiceProvider

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()
            SuspendLayout()

            Text = "Property Page Designer View"    ' For Debug

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

            ' Scale the width of the Configuration/Platform combo boxes
            ConfigurationComboBox.Width = DpiAwareness.LogicalToDeviceUnits(Handle, ConfigurationComboBox.Width)
            PlatformComboBox.Width = DpiAwareness.LogicalToDeviceUnits(Handle, PlatformComboBox.Width)

            'Start out with the assumption that the configuration/platform comboboxes
            '  are invisible, otherwise they will flicker visible before being turned off.
            ConfigurationPanel.Visible = False

            ResumeLayout(False)
            PerformLayout()
        End Sub

        Public WithEvents ConfigDividerLine As Label
        Public WithEvents PlatformComboBox As ComboBox
        Public WithEvents PlatformLabel As Label
        Public WithEvents PropPageDesignerViewLayoutPanel As TableLayoutPanel
        Public WithEvents ConfigurationFlowLayoutPanel As FlowLayoutPanel
        Public WithEvents ConfigurationTableLayoutPanel As TableLayoutPanel
        Public WithEvents PLatformTableLayoutPanel As TableLayoutPanel
        Public WithEvents ConfigurationPanel As TableLayoutPanel

        'Required by the Windows Form Designer
        Private ReadOnly _components As IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <DebuggerNonUserCode()> Private Sub InitializeComponent()
            Dim resources As ComponentResourceManager = New ComponentResourceManager(GetType(PropPageDesignerView))
            ConfigurationComboBox = New ComboBox
            PlatformLabel = New Label
            PlatformComboBox = New ComboBox
            ConfigDividerLine = New Label
            ConfigurationLabel = New Label
            PropertyPagePanel = New ScrollablePanel
            PropPageDesignerViewLayoutPanel = New TableLayoutPanel
            ConfigurationPanel = New TableLayoutPanel
            ConfigurationFlowLayoutPanel = New FlowLayoutPanel
            ConfigurationTableLayoutPanel = New TableLayoutPanel
            PLatformTableLayoutPanel = New TableLayoutPanel
            PropPageDesignerViewLayoutPanel.SuspendLayout()
            ConfigurationPanel.SuspendLayout()
            ConfigurationFlowLayoutPanel.SuspendLayout()
            ConfigurationTableLayoutPanel.SuspendLayout()
            PLatformTableLayoutPanel.SuspendLayout()
            SuspendLayout()
            '
            'ConfigurationComboBox
            '
            resources.ApplyResources(ConfigurationComboBox, "ConfigurationComboBox")
            ConfigurationComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            ConfigurationComboBox.FormattingEnabled = True
            ConfigurationComboBox.Name = "ConfigurationComboBox"
            '
            'PlatformLabel
            '
            resources.ApplyResources(PlatformLabel, "PlatformLabel")
            PlatformLabel.Name = "PlatformLabel"
            '
            'PlatformComboBox
            '
            resources.ApplyResources(PlatformComboBox, "PlatformComboBox")
            PlatformComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            PlatformComboBox.FormattingEnabled = True
            PlatformComboBox.Name = "PlatformComboBox"
            '
            'ConfigDividerLine
            '
            resources.ApplyResources(ConfigDividerLine, "ConfigDividerLine")
            ConfigDividerLine.BorderStyle = BorderStyle.Fixed3D
            ConfigDividerLine.Name = "ConfigDividerLine"
            '
            'ConfigurationLabel
            '
            resources.ApplyResources(ConfigurationLabel, "ConfigurationLabel")
            ConfigurationLabel.Name = "ConfigurationLabel"
            '
            'PropertyPagePanel
            '
            resources.ApplyResources(PropertyPagePanel, "PropertyPagePanel")
            PropertyPagePanel.Name = "PropertyPagePanel"
            '
            'PropPageDesignerViewLayoutPanel
            '
            resources.ApplyResources(PropPageDesignerViewLayoutPanel, "PropPageDesignerViewLayoutPanel")
            PropPageDesignerViewLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            PropPageDesignerViewLayoutPanel.Controls.Add(ConfigurationPanel, 0, 0)
            PropPageDesignerViewLayoutPanel.Controls.Add(PropertyPagePanel, 0, 1)
            PropPageDesignerViewLayoutPanel.Name = "PropPageDesignerViewLayoutPanel"
            PropPageDesignerViewLayoutPanel.RowStyles.Add(New RowStyle)
            PropPageDesignerViewLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
            '
            'ConfigurationPanel
            '
            resources.ApplyResources(ConfigurationPanel, "ConfigurationPanel")
            ConfigurationPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            ConfigurationPanel.Controls.Add(ConfigDividerLine, 1, 1)
            ConfigurationPanel.Controls.Add(ConfigurationFlowLayoutPanel, 0, 0)
            ConfigurationPanel.Name = "ConfigurationPanel"
            ConfigurationPanel.RowStyles.Add(New RowStyle)
            ConfigurationPanel.RowStyles.Add(New RowStyle)
            '
            'ConfigurationFlowLayoutPanel
            '
            resources.ApplyResources(ConfigurationFlowLayoutPanel, "ConfigurationFlowLayoutPanel")
            ConfigurationFlowLayoutPanel.CausesValidation = False
            ConfigurationFlowLayoutPanel.Controls.Add(ConfigurationTableLayoutPanel)
            ConfigurationFlowLayoutPanel.Controls.Add(PLatformTableLayoutPanel)
            ConfigurationFlowLayoutPanel.Name = "ConfigurationFlowLayoutPanel"
            '
            'ConfigurationTableLayoutPanel
            '
            resources.ApplyResources(ConfigurationTableLayoutPanel, "ConfigurationTableLayoutPanel")
            ConfigurationTableLayoutPanel.ColumnStyles.Add(New ColumnStyle)
            ConfigurationTableLayoutPanel.ColumnStyles.Add(New ColumnStyle)
            ConfigurationTableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 10.0!))
            ConfigurationTableLayoutPanel.Controls.Add(ConfigurationComboBox, 1, 0)
            ConfigurationTableLayoutPanel.Controls.Add(ConfigurationLabel, 0, 0)
            ConfigurationTableLayoutPanel.Name = "ConfigurationTableLayoutPanel"
            ConfigurationTableLayoutPanel.RowStyles.Add(New RowStyle)
            '
            'PLatformTableLayoutPanel
            '
            resources.ApplyResources(PLatformTableLayoutPanel, "PLatformTableLayoutPanel")
            PLatformTableLayoutPanel.ColumnStyles.Add(New ColumnStyle)
            PLatformTableLayoutPanel.ColumnStyles.Add(New ColumnStyle)
            PLatformTableLayoutPanel.Controls.Add(PlatformLabel, 0, 0)
            PLatformTableLayoutPanel.Controls.Add(PlatformComboBox, 1, 0)
            PLatformTableLayoutPanel.Name = "PLatformTableLayoutPanel"
            PLatformTableLayoutPanel.RowStyles.Add(New RowStyle)
            '
            'PropPageDesignerView
            '
            resources.ApplyResources(Me, "$this")
            Controls.Add(PropPageDesignerViewLayoutPanel)
            Name = "PropPageDesignerView"
            PropPageDesignerViewLayoutPanel.ResumeLayout(False)
            PropPageDesignerViewLayoutPanel.PerformLayout()
            ConfigurationPanel.ResumeLayout(False)
            ConfigurationPanel.PerformLayout()
            ConfigurationFlowLayoutPanel.ResumeLayout(False)
            ConfigurationFlowLayoutPanel.PerformLayout()
            ConfigurationTableLayoutPanel.ResumeLayout(False)
            ConfigurationTableLayoutPanel.PerformLayout()
            PLatformTableLayoutPanel.ResumeLayout(False)
            PLatformTableLayoutPanel.PerformLayout()
            ResumeLayout(False)

        End Sub

#End Region

        'The currently-loaded page and its site
        Private _loadedPage As OleInterop.IPropertyPage
        Private _loadedPageSite As PropertyPageSite

        'True once we have been initialized completely.
        Private _fInitialized As Boolean

        'If true, we ignore the selected index changed event
        Private _ignoreSelectedIndexChanged As Boolean

        Private _errorControl As Control 'Displayed error control, if any

        Public Const SW_HIDE As Integer = 0
        Public Const SW_SHOWNORMAL As Integer = 1
        Public Const SW_SHOW As Integer = 5

        Public WithEvents ConfigurationLabel As Label
        Public WithEvents PropertyPagePanel As ScrollablePanel
        Public WithEvents ConfigurationComboBox As ComboBox

        Private _rootDesigner As PropPageDesignerRootDesigner
        Private _projectHierarchy As IVsHierarchy

        Private _uiShellService As IVsUIShell
        Private _uiShell5Service As IVsUIShell5

        ' The ConfigurationState object from the project designer.  This is shared among all the prop page designers
        '   for this project designer.
        Private _configurationState As ConfigurationState

        'True if we should check for simplified config mode having changed (used to keep from checking multiple times in a row)
        Private _needToCheckForModeChanges As Boolean

        'The number of undo units that were available when the page was in a clean state.
        Private _undoUnitsOnStackAtCleanState As Integer

        ' The UndoEngine for this designer
        Private WithEvents _undoEngine As UndoEngine

        ' The DesignerHost for this designer
        Private WithEvents _designerHost As IDesignerHost

        'True iff the property page is currently activated
        Private _isPageActivated As Boolean

        'True iff the property page is hosted through native SetParent and not as a Windows Form child control
        Private _isNativeHostedPropertyPage As Boolean

        'Listen for font/color changes from the shell
        Private WithEvents _broadcastMessageEventsHelper As ShellUtil.BroadcastMessageEventsHelper

#Region "Constructor"

        ''' <summary>
        ''' View constructor 
        ''' </summary>
        ''' <param name="RootDesigner"></param>
        Public Sub New(RootDesigner As PropPageDesignerRootDesigner)
            Me.New()
            SetSite(RootDesigner)
        End Sub

#End Region

#Region "Dispose/IDisposable"
        'UserControl overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                Try
                    If _rootDesigner IsNot Nothing Then
                        _rootDesigner.RemoveMenuCommands()
                    End If

                    _undoEngine = Nothing
                    UnLoadPage()
                    If _components IsNot Nothing Then
                        _components.Dispose()
                    End If
                    _configurationState = Nothing
                Catch ex As Exception When ReportWithoutCrash(ex, NameOf(Dispose), NameOf(PropPageDesignerView))
                    'Don't throw here trying to cleanup
                End Try
            End If
            MyBase.Dispose(disposing)
        End Sub
#End Region

        ''' <summary>
        ''' Get DesignerHost
        ''' </summary>
        Public ReadOnly Property DesignerHost As IDesignerHost
            Get
                Return TryCast(GetService(GetType(IDesignerHost)), IDesignerHost)
            End Get
        End Property

        ''' <summary>
        ''' Property page we host
        ''' </summary>
        Public ReadOnly Property PropPage As OleInterop.IPropertyPage
            Get
                Return _loadedPage
            End Get
        End Property

        Private _isConfigPage As Boolean
        Public Property IsConfigPage As Boolean
            Get
                Return _isConfigPage
            End Get
            Set
                _isConfigPage = Value
            End Set
        End Property

        ''' <summary>
        ''' Retrieves the IVsUIShell service
        ''' </summary>
        Private ReadOnly Property VsUIShellService As IVsUIShell
            Get
                If _uiShellService Is Nothing Then
                    If VBPackageInstance IsNot Nothing Then
                        _uiShellService = TryCast(VBPackageInstance.GetService(GetType(IVsUIShell)), IVsUIShell)
                    Else
                        _uiShellService = TryCast(GetService(GetType(IVsUIShell)), IVsUIShell)
                    End If
                End If

                Return _uiShellService
            End Get
        End Property

        ''' <summary>
        ''' Retrieves the IVsUIShell5 service
        ''' </summary>
        Private ReadOnly Property VsUIShell5Service As IVsUIShell5
            Get
                If _uiShell5Service Is Nothing Then
                    Dim VsUiShell As IVsUIShell = VsUIShellService
                    If VsUiShell IsNot Nothing Then
                        _uiShell5Service = TryCast(VsUiShell, IVsUIShell5)
                    End If
                End If

                Return _uiShell5Service
            End Get
        End Property

        Private _dteProject As EnvDTE.Project

        Public ReadOnly Property DTEProject As EnvDTE.Project
            Get
                Return _dteProject
            End Get
        End Property

        ''' <summary>
        ''' True iff the property page is hosted through native SetParent and not as a Windows Form child control.
        ''' Returns False if the property page is not currently activated
        ''' </summary>
        Public ReadOnly Property IsNativeHostedPropertyPageActivated As Boolean
            Get
                Return _isPageActivated AndAlso _isNativeHostedPropertyPage
            End Get
        End Property

        ''' <summary>
        ''' Gets the browse object for the project.  This is what is passed to SetObjects for
        '''   non-config-dependent pages
        ''' </summary>
        Private Function GetProjectBrowseObject() As Object
            Dim BrowseObject As Object = Nothing
            VSErrorHandler.ThrowOnFailure(_projectHierarchy.GetProperty(VSITEMID.ROOT, __VSHPROPID.VSHPROPID_BrowseObject, BrowseObject))
            Return BrowseObject
        End Function

        Private _vsCfgProvider As IVsCfgProvider2
        Private ReadOnly Property VsCfgProvider As IVsCfgProvider2
            Get
                If _vsCfgProvider Is Nothing Then
                    Dim Value As Object = Nothing

                    VSErrorHandler.ThrowOnFailure(_projectHierarchy.GetProperty(VSITEMID.ROOT, __VSHPROPID.VSHPROPID_ConfigurationProvider, Value))

                    _vsCfgProvider = CType(Value, IVsCfgProvider2)
                End If
                Return _vsCfgProvider
            End Get
        End Property

        ''' <summary>
        ''' Initialization routine called by the ApplicationDesignerView when the page is first activated
        ''' </summary>
        ''' <param name="DTEProject"></param>
        ''' <param name="PropPage"></param>
        ''' <param name="PropPageSite"></param>
        ''' <param name="Hierarchy"></param>
        ''' <param name="IsConfigPage"></param>
        Public Sub Init(DTEProject As EnvDTE.Project, PropPage As OleInterop.IPropertyPage, PropPageSite As PropertyPageSite, Hierarchy As IVsHierarchy, IsConfigPage As Boolean)
            Debug.Assert(_dteProject Is Nothing, "Init() called twice?")

            Debug.Assert(DTEProject IsNot Nothing, "DTEProject is Nothing")
            Debug.Assert(PropPage IsNot Nothing)
            Debug.Assert(PropPageSite IsNot Nothing)
            Debug.Assert(Hierarchy IsNot Nothing)

            _dteProject = DTEProject
            _loadedPage = PropPage
            _loadedPageSite = PropPageSite
            _projectHierarchy = Hierarchy
            Me.IsConfigPage = IsConfigPage

            SuspendLayout()
            ConfigurationPanel.SuspendLayout()
            PropertyPagePanel.SuspendLayout()

            SetDialogFont()

            Dim menuCommands As New ArrayList()
            Dim cutCmd As New DesignerMenuCommand(_rootDesigner, Constants.MenuConstants.CommandIDVSStd97cmdidCut, AddressOf DisabledMenuCommandHandler) With {
                .Enabled = False
            }
            menuCommands.Add(cutCmd)

            _rootDesigner.RegisterMenuCommands(menuCommands)

            ' Get the ConfigurationState object from the project designer
            _configurationState = DirectCast(_loadedPageSite.GetService(GetType(ConfigurationState)), ConfigurationState)
            If _configurationState Is Nothing Then
                Debug.Fail("Couldn't get ConfigurationState service")
                Throw New Package.InternalException
            End If
            If IsConfigPage Then
                AddHandler _configurationState.SelectedConfigurationChanged, AddressOf ConfigurationState_SelectedConfigurationChanged
                AddHandler _configurationState.ConfigurationListAndSelectionChanged, AddressOf ConfigurationState_ConfigurationListAndSelectionChanged

                'Note: we only hook this up for config pages because the situations where we (currently) need to clear the undo/redo stack only
                '  affects config pages (when a config/platform is deleted or renamed).
                AddHandler _configurationState.ClearConfigPageUndoRedoStacks, AddressOf ConfigurationState_ClearConfigPageUndoRedoStacks
            End If

            'This notification is needed by config and non-config pages
            AddHandler _configurationState.SimplifiedConfigModeChanged, AddressOf ConfigurationState_SimplifiedConfigModeChanged

            'Scale the comboboxes widths if necessary, for High-DPI
            ConfigurationComboBox.Size = DpiAwareness.LogicalToDeviceSize(Handle, ConfigurationComboBox.Size)
            PlatformComboBox.Size = DpiAwareness.LogicalToDeviceSize(Handle, PlatformComboBox.Size)

            'Set up configuration/platform comboboxes
            SetConfigDropdownVisibility()
            UpdateConfigLists() 'This is done initially for config and non-config pages

            'Set the initial dropdown selections
            If IsConfigPage Then
                ChangeSelectedComboBoxIndicesWithoutNotification(_configurationState.SelectedConfigIndex, _configurationState.SelectedPlatformIndex)
            End If

            'Populate the page initially
            SetObjectsForSelectedConfigs()

            ActivatePage(PropPage)

            ConfigurationPanel.ResumeLayout(True)
            PropertyPagePanel.ResumeLayout(True)
            ResumeLayout(True)

            'PERF: no need to call UpdatePageSize here - Activate() is already passed in a rectangle to 
            '  move the control to initially
            'UpdatePageSize() 

            _needToCheckForModeChanges = False
            _fInitialized = True

            If _undoEngine Is Nothing Then
                _undoEngine = DirectCast(GetService(GetType(UndoEngine)), UndoEngine)
            End If

            If _designerHost Is Nothing Then
                _designerHost = DirectCast(GetService(GetType(IDesignerHost)), IDesignerHost)
            End If
        End Sub

        ''' <summary>
        ''' Occurs after an undo or redo operation has completed.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub OnUndoEngineUndone(sender As Object, e As EventArgs) Handles _undoEngine.Undone
            'Tell the project designer it needs to refresh its dirty status
            If _loadedPageSite IsNot Nothing Then
                Dim AppDesignerView As ApplicationDesignerView = TryCast(_loadedPageSite.GetService(GetType(ApplicationDesignerView)), ApplicationDesignerView)
                If AppDesignerView IsNot Nothing Then
                    AppDesignerView.DelayRefreshDirtyIndicators()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Update dirty state of the appdesigner (if any) after a transaction is closed
        '''  Sometimes the we may have tried to update the dirty state when a transaction is open (i.e. when opening files
        '''  from SCC). It is not possible for us to update the dirty state while a transaction is active, so we have to wait
        '''  until the transaction is closed. 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub OnDesignerHostTransactionClosed(sender As Object, e As DesignerTransactionCloseEventArgs) Handles _designerHost.TransactionClosed
            If _loadedPageSite IsNot Nothing Then
                Dim AppDesignerView As ApplicationDesignerView = TryCast(_loadedPageSite.GetService(GetType(ApplicationDesignerView)), ApplicationDesignerView)
                If AppDesignerView IsNot Nothing Then
                    AppDesignerView.DelayRefreshDirtyIndicators()
                End If
            End If
        End Sub

        Private Sub OnThemeChanged()
            If TypeOf PropPage Is PropPageBase AndAlso CType(PropPage, PropPageBase).SupportsTheming Then
                BackColor = ShellUtil.GetProjectDesignerThemeColor(VsUIShell5Service, "Background", __THEMEDCOLORTYPE.TCT_Background, SystemColors.Control)
            Else
                BackColor = SystemColors.Control
            End If

            ConfigurationPanel.BackColor = BackColor
            PropertyPagePanel.BackColor = BackColor
        End Sub

        ''' <summary>
        ''' We've gotta tell the renderer whenever the system colors change...
        ''' </summary>
        ''' <param name="msg"></param>
        ''' <param name="wparam"></param>
        ''' <param name="lparam"></param>
        Private Sub OnBroadcastMessageEventsHelperBroadcastMessage(msg As UInteger, wParam As IntPtr, lParam As IntPtr) Handles _broadcastMessageEventsHelper.BroadcastMessage
            Select Case msg
                Case Win32Constant.WM_PALETTECHANGED, Win32Constant.WM_SYSCOLORCHANGE, Win32Constant.WM_THEMECHANGED
                    OnThemeChanged()
                Case Win32Constant.WM_SETTINGCHANGE
                    Dim newFont As Font = ShellUtil.FontChangeMonitor.GetDialogFont(Me)
                    If Not newFont.Equals(Font) Then
                        Font = newFont
                    End If
            End Select
        End Sub

        Friend Sub SetControls(firstControl As Boolean)
            If firstControl Then
                If _isNativeHostedPropertyPage Then
                    'Try to set initial focus to the property page, not the configuration panel
                    FocusFirstOrLastPropertyPageControl(True)
                Else
                    ' Select the configuration panel to ensure it gains focus. For configuration pages, this
                    ' ensures that the configuration panel receives focus and allows a screen reader to give
                    ' the page context before reading the values of any properties themselves. For other pages
                    ' this ensures that the first control of the page receives focus, which allows tab navigation
                    ' to work in a reliable and predicable manner for users who can only use the keyboard.
                    SelectNextControl(ConfigurationPanel, forward:=True, tabStopOnly:=True, nested:=True, wrap:=True)
                End If
            Else
                FocusFirstOrLastPropertyPageControl(False)
            End If
        End Sub

        ''' <summary>
        ''' Show the property page 
        ''' </summary>
        ''' <param name="PropPage"></param>
        Public Sub ActivatePage(PropPage As OleInterop.IPropertyPage)
            Switches.TracePDPerfBegin("PropPageDesignerView.ActivatePage")
            If PropPage Is Nothing Then
                'Property page failed to load - just give empty page
            Else
                'Set the Undo site before activation
                Dim vsProjectDesignerPage = TryCast(PropPage, IVsProjectDesignerPage)
                If vsProjectDesignerPage IsNot Nothing Then
                    vsProjectDesignerPage.SetSite(Me)
                End If

                'Activate the page
                Debug.Assert(Not Handle.Equals(IntPtr.Zero), "Window not yet created")
                'Force creation of the control to get hwnd 
                If Handle.Equals(IntPtr.Zero) Then
                    CreateControl()
                End If

                Try
                    ' Check the minimum size for the control and make sure that we show scrollbars
                    ' if the PropertyPagePanel becomes smaller...
                    Dim Info As OleInterop.PROPPAGEINFO() = New OleInterop.PROPPAGEINFO(0) {}
                    If PropPage IsNot Nothing Then
                        PropPage.GetPageInfo(Info)
                        PropertyPagePanel.AutoScrollMinSize = New Size(Info(0).SIZE.cx + Padding.Right + Padding.Left, Info(0).SIZE.cy + Padding.Top + Padding.Bottom)
                    End If

                    PropPage.Activate(PropertyPagePanel.Handle, New OleInterop.RECT() {GetPageRect()}, 0)

                    PropPage.Show(SW_SHOW)
                    'UpdateWindowStyles(Me.Handle)
                    'Me.MinimumSize = GetMaxSize()

                    ' Dev10 Bug 905047
                    ' Explicitly initialize the UI cue state so that focus and keyboard cues work.
                    ' We need to do this explicitly since this UI isn't a dialog (where the state
                    ' would have been automatically initialized)
                    InitializeStateOfUICues()

                    ' It is a managed control, we should update AutoScrollMinSize
                    If PropertyPagePanel.Controls.Count > 0 Then
                        Dim controlSize As Size = PropertyPagePanel.Controls(0).Size
                        PropertyPagePanel.AutoScrollMinSize = New Size(
                            Math.Min(controlSize.Width + Padding.Right + Padding.Left, PropertyPagePanel.AutoScrollMinSize.Width),
                            Math.Min(controlSize.Height + Padding.Top + Padding.Bottom, PropertyPagePanel.AutoScrollMinSize.Height))
                    End If

                    _isPageActivated = True
                    OnThemeChanged()

                    'Is the control hosted natively via SetParent?
                    If PropertyPagePanel.Controls.Count > 0 Then
                        _isNativeHostedPropertyPage = False
                    Else
                        _isNativeHostedPropertyPage = True
                    End If

                    If _isNativeHostedPropertyPage Then
                        'Try to set initial focus to the property page, not the configuration panel
                        FocusFirstOrLastPropertyPageControl(True)
                    Else
                        ' Select the configuration panel to ensure it gains focus. For configuration pages, this
                        ' ensures that the configuration panel receives focus and allows a screen reader to give
                        ' the page context before reading the values of any properties themselves. For other pages
                        ' this ensures that the first control of the page receives focus, which allows tab navigation
                        ' to work in a reliable and predicable manner for users who can only use the keyboard.
                        SelectNextControl(ConfigurationPanel, forward:=True, tabStopOnly:=True, nested:=True, wrap:=True)
                    End If

                    ''Only set the undo redo state if we are loading a new page
                    SetUndoRedoCleanState()
                Catch ex As Exception When ReportWithoutCrash(ex, NameOf(ActivatePage), NameOf(PropPageDesignerView))
                    'There was a problem displaying the property page.  Show the error control.
                    DisplayErrorControl(ex)
                End Try
            End If
            Switches.TracePDPerfEnd("PropPageDesignerView.ActivatePage")
        End Sub

        ''' <summary>
        ''' Display the error control instead of a property page
        ''' </summary>
        ''' <param name="ex">The exception to retrieve the error message from</param>
        Private Sub DisplayErrorControl(ex As Exception)

            UnLoadPage()
            PropertyPagePanel.SuspendLayout()
            ConfigurationPanel.Visible = False
            If TypeOf ex Is PropertyPageException AndAlso Not DirectCast(ex, PropertyPageException).ShowHeaderAndFooterInErrorControl Then
                _errorControl = New ErrorControl(ex.Message)
            Else
                _errorControl = New ErrorControl(My.Resources.Designer.APPDES_ErrorLoadingPropPage & vbCrLf & DebugMessageFromException(ex))
            End If
            _errorControl.Dock = DockStyle.Fill
            _errorControl.Visible = True
            PropertyPagePanel.Controls.Add(_errorControl)
            PropertyPagePanel.ResumeLayout(True)
        End Sub

        ''' <summary>
        ''' Hide and deactivate the property page
        ''' </summary>
        Public Sub UnLoadPage()
            _isPageActivated = False
            _isNativeHostedPropertyPage = False

            If _loadedPage IsNot Nothing Then
                'Store in local and clear member first in case of throw by deactivate
                Dim Page As OleInterop.IPropertyPage = _loadedPage
                _loadedPage = Nothing
                Try
                    Page.SetObjects(0, Nothing)
                    Page.Deactivate()
                Catch ex As Exception When ReportWithoutCrash(ex, NameOf(UnLoadPage), NameOf(PropPageDesignerView))
                End Try
            End If
            If _errorControl IsNot Nothing Then
                Controls.Remove(_errorControl)
                _errorControl.Dispose()
                _errorControl = Nothing
            End If

            If _configurationState IsNot Nothing Then
                RemoveHandler _configurationState.SelectedConfigurationChanged, AddressOf ConfigurationState_SelectedConfigurationChanged
                RemoveHandler _configurationState.ConfigurationListAndSelectionChanged, AddressOf ConfigurationState_ConfigurationListAndSelectionChanged
                RemoveHandler _configurationState.ClearConfigPageUndoRedoStacks, AddressOf ConfigurationState_ClearConfigPageUndoRedoStacks
                RemoveHandler _configurationState.SimplifiedConfigModeChanged, AddressOf ConfigurationState_SimplifiedConfigModeChanged
            End If
        End Sub

        ''' <summary>
        ''' Our site - always of type PropPageDesignerRootDesigner
        ''' </summary>
        ''' <param name="RootDesigner"></param>
        Private Sub SetSite(RootDesigner As PropPageDesignerRootDesigner) 'Implements OLE.Interop.IObjectWithSite.SetSite
            _rootDesigner = RootDesigner
            _broadcastMessageEventsHelper = New ShellUtil.BroadcastMessageEventsHelper(Me)
            OnThemeChanged()
        End Sub

        ''' <summary>
        ''' GetService helper
        ''' </summary>
        ''' <param name="ServiceType"></param>
        Public Shadows Function GetService(ServiceType As Type) As Object Implements IServiceProvider.GetService
            Dim Service As Object
            Service = _rootDesigner.GetService(ServiceType)
            Return Service
        End Function

        ''' <summary>
        ''' Get the size of the hosting client rect for sizing the property page 
        ''' </summary>
        Private Function GetPageRect() As OleInterop.RECT
            Dim ClientRect As New OleInterop.RECT
            ' We should use DisplayRectangle.Left/Top here, so the child page could work with auto-scroll
            With ClientRect
                .left = PropertyPagePanel.DisplayRectangle.Left
                .top = PropertyPagePanel.DisplayRectangle.Top
                .right = PropertyPagePanel.ClientSize.Width + .left
                .bottom = PropertyPagePanel.ClientSize.Height + .top
            End With
            Return ClientRect
        End Function

        ''' <summary>
        ''' Update the hosted property page size
        ''' </summary>
        Private Sub UpdatePageSize()
            If _loadedPage IsNot Nothing Then
                Dim RectArray As OleInterop.RECT() = New OleInterop.RECT() {GetPageRect()}
                Switches.TracePDPerfBegin("PropPageDesignerView.UpdatePageSize (" _
                    & RectArray(0).right - RectArray(0).left & ", " & RectArray(0).bottom - RectArray(0).top & ")")
                _loadedPage.Move(RectArray)
                Switches.TracePDPerfEnd("PropPageDesignerView.UpdatePageSize")
            End If
        End Sub

        Protected Overrides Sub OnLayout(e As LayoutEventArgs)
            Switches.TracePDPerfBegin(e, "PropPageDesignerView.OnLayout()")
            MyBase.OnLayout(e)

            ' Hard coded to change the size of the LayoutPanel to fit our clientSize. Otherwise, it will pick its own size...
            If DisplayRectangle <> Rectangle.Empty Then
                PropPageDesignerViewLayoutPanel.Size = ClientSize
                UpdatePageSize()
            End If
            Switches.TracePDPerfEnd("PropPageDesignerView.OnLayout()")
        End Sub

#Region "IVsProjectDesignerPageSite"

        ''' <summary>
        ''' This is part of the undo host code for the property page.  
        ''' We pass this interface to the property pages implementation of IVsProjectDesignerPage.SetSite
        ''' </summary>
        Private Sub OnPropertyChanged(Component As Component, PropDesc As PropertyDescriptor, OldValue As Object, NewValue As Object)
            Dim ChangeService As IComponentChangeService = DirectCast(GetService(GetType(IComponentChangeService)), IComponentChangeService)
            ChangeService.OnComponentChanged(Component, PropDesc, OldValue, NewValue)
        End Sub

        ''' <summary>
        ''' If a property page hosted by the Project Designer wants to support automatic Undo/Redo, it must call
        '''   call this method on the IVsProjectDesignerPageSite after a property value is changed.
        ''' </summary>
        ''' <param name="propertyName">The name of the property whose value has changed.</param>
        ''' <param name="propertyDescriptor">A PropertyDescriptor that describes the given property.</param>
        ''' <param name="oldValue">The previous value of the property.</param>
        ''' <param name="newValue">The new value of the property.</param>
        Public Sub IVsProjectDesignerPageSite_OnPropertyChanged(PropertyName As String, PropertyDescriptor As PropertyDescriptor, OldValue As Object, NewValue As Object) Implements IVsProjectDesignerPageSite.OnPropertyChanged
            'Note: we wrap the property descriptor here because it allows us to intercept the GetValue/SetValue calls and therefore
            '  more finely control the undo/redo process.
            OnPropertyChanged(_rootDesigner.Component, New PropertyPagePropertyDescriptor(PropertyDescriptor, PropertyName), OldValue, NewValue)
        End Sub

        ''' <summary>
        ''' This is part of the undo host code for the property page.  
        ''' We pass this interface to the property pages implementation of IVsProjectDesignerPage.SetSite
        ''' </summary>
        Private Sub OnPropertyChanging(Component As Component, PropDesc As PropertyDescriptor)
            Dim ChangeService As IComponentChangeService = DirectCast(GetService(GetType(IComponentChangeService)), IComponentChangeService)

            If PropDesc Is Nothing Then
                Debug.Fail("We should not be here")
                ChangeService.OnComponentChanging(Component, Nothing)
            Else
                ChangeService.OnComponentChanging(Component, PropDesc)
            End If
        End Sub

        ''' <summary>
        ''' If a property page hosted by the Project Designer wants to support automatic Undo/Redo, it must call
        '''   call this method on the IVsProjectDesignerPageSite before a property value is changed.  This allows 
        '''   the site to query for the current value of the property and save it for later use in handling Undo/Redo.
        ''' </summary>
        ''' <param name="propertyName">The name of the property whose value is about to change.</param>
        ''' <param name="propertyDescriptor">A PropertyDescriptor that describes the given property.</param>
        Public Sub IVsProjectDesignerPageSite_OnPropertyChanging(PropertyName As String, PropertyDescriptor As PropertyDescriptor) Implements IVsProjectDesignerPageSite.OnPropertyChanging
            'Note: we wrap the property descriptor here because it allows us to intercept the GetValue/SetValue calls and therefore
            '  more finely control the undo/redo process.
            OnPropertyChanging(_rootDesigner.Component, New PropertyPagePropertyDescriptor(PropertyDescriptor, PropertyName))
        End Sub

        ''' <summary>
        ''' Retrieves a transaction which can be used to group multiple property changes into a single transaction, so that
        '''   they appear to the user as a single Undo/Redo unit.  The transaction must be committed or cancelled after the
        '''   property changes are made.
        ''' </summary>
        ''' <param name="description">The localized description string to use for the transaction.  This will appear as the
        '''   description for the Undo/Redo unit.</param>
        Public Function GetTransaction(Description As String) As DesignerTransaction Implements IVsProjectDesignerPageSite.GetTransaction
            Dim DesignerHost As IDesignerHost
            DesignerHost = DirectCast(GetService(GetType(IDesignerHost)), IDesignerHost)
            Return DesignerHost.CreateTransaction(Description)
        End Function

#End Region

        ''' <summary>
        ''' Set font for controls on the Configuration panel
        ''' </summary>
        Private Sub SetDialogFont()
            Font = GetDialogFont()
        End Sub

        ''' <summary>
        ''' Pick font to use in this dialog page
        ''' </summary>
        Private ReadOnly Property GetDialogFont As Font
            Get
                Dim uiSvc As IUIService = CType(GetService(GetType(IUIService)), IUIService)
                If uiSvc IsNot Nothing Then
                    Return CType(uiSvc.Styles("DialogFont"), Font)
                End If

                Debug.Fail("Couldn't get a IUIService... cheating instead :)")

                Return DefaultFont
            End Get
        End Property

        'Standard title for messageboxes, etc.
        Private ReadOnly _messageBoxCaption As String = My.Resources.Designer.APPDES_Title

        ''' <summary>
        ''' Displays a message box using the Visual Studio-approved manner.
        ''' </summary>
        ''' <param name="Message">The message text.</param>
        ''' <param name="Buttons">Which buttons to show</param>
        ''' <param name="Icon">the icon to show</param>
        ''' <param name="DefaultButton">Which button should be default?</param>
        ''' <param name="HelpLink">The help link</param>
        ''' <returns>One of the DialogResult values</returns>
        Public Function DsMsgBox(Message As String,
                Buttons As MessageBoxButtons,
                Icon As MessageBoxIcon,
                Optional DefaultButton As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1,
                Optional HelpLink As String = Nothing) As DialogResult

            Return DesignerMessageBox.Show(_rootDesigner, Message, _messageBoxCaption,
                Buttons, Icon, DefaultButton, HelpLink)
        End Function

        ''' <summary>
        ''' Displays a designer error message
        ''' </summary>
        ''' <param name="Message"></param>
        Public Sub ShowErrorMessage(Message As String, Optional HelpLink As String = Nothing)
            DsMsgBox(Message, MessageBoxButtons.OK, MessageBoxIcon.Error, HelpLink:=HelpLink)
        End Sub

#Region "Configuration/Platform Comboboxes and related code"

        ''' <summary>
        ''' Sets whether or not the configuration/platform dropdowns are visible
        ''' </summary>
        Private Sub SetConfigDropdownVisibility()
            ConfigurationPanel.Visible = Not _configurationState.IsSimplifiedConfigMode()

            If IsConfigPage Then
                ConfigurationPanel.Enabled = True
            Else
                'Non-configuration pages should have the configuration panel visible but disabled, and the text should be "N/A"
                ConfigurationPanel.Enabled = False

                ConfigurationComboBox.Items.Add(My.Resources.Designer.PPG_NotApplicable)
                ConfigurationComboBox.SelectedIndex = 0
                PlatformComboBox.Items.Add(My.Resources.Designer.PPG_NotApplicable)
                PlatformComboBox.SelectedIndex = 0
            End If

            'Update layout with the change in visibility
            PerformLayout()
        End Sub

        ''' <summary>
        ''' Changes the indices of the configuration and platform comboboxes, without causing a notify to be sent to the
        '''   ConfigurationState
        ''' </summary>
        ''' <param name="NewSelectedConfigIndex">New index into the ConfigurationComboBox</param>
        ''' <param name="NewSelectedPlatformIndex">New index into the PlatformComboBox</param>
        Private Sub ChangeSelectedComboBoxIndicesWithoutNotification(NewSelectedConfigIndex As Integer, NewSelectedPlatformIndex As Integer)
            Debug.Assert(Not _ignoreSelectedIndexChanged)

            Dim OldIgnoreSelectedIndexChanged As Boolean = _ignoreSelectedIndexChanged
            _ignoreSelectedIndexChanged = True
            Try
                ConfigurationComboBox.SelectedIndex = NewSelectedConfigIndex
                PlatformComboBox.SelectedIndex = NewSelectedPlatformIndex
            Finally
                _ignoreSelectedIndexChanged = OldIgnoreSelectedIndexChanged
            End Try
        End Sub

        ''' <summary>
        ''' Occurs when the user selects an item in the configuration combobox or the platform combobox
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub SelectedConfigurationOrPlatformIndexChanged(sender As Object, e As EventArgs) _
                        Handles ConfigurationComboBox.SelectedIndexChanged, PlatformComboBox.SelectedIndexChanged, PlatformComboBox.SelectedIndexChanged

            If _fInitialized AndAlso Not _ignoreSelectedIndexChanged Then
                Debug.Assert(IsConfigPage)
                If IsConfigPage Then
                    'Notify the ConfigurationState of the change.  It will in turn notify us via SelectedConfigurationChanged
                    _configurationState.ChangeSelection(ConfigurationComboBox.SelectedIndex, PlatformComboBox.SelectedIndex, FireNotifications:=True)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Fired when the selected configuration is changed on another property page or in the
        '''   configuration manager.
        ''' </summary>
        ''' <remarks>
        ''' Listener must update their selection state by querying SelectedConfigIndex 
        '''   and SelectedPlatformIndex.
        ''' </remarks>>
        Private Sub ConfigurationState_SelectedConfigurationChanged()
            Debug.Assert(IsConfigPage)
            If IsConfigPage Then
                'Update combobox selections
                ChangeSelectedComboBoxIndicesWithoutNotification(_configurationState.SelectedConfigIndex, _configurationState.SelectedPlatformIndex)

                '... and tell the page to update based on the new selection 'CONSIDER delaying this call until we're the active designer
                SetObjectsForSelectedConfigs()
            End If
        End Sub

        ''' <summary>
        ''' Raised when the configuration/platform lists have changed.  Note that this
        '''   will *not* be followed by a SelectedConfigurationChanged event, but the
        '''   listener should still update the selection as well as their lists.
        ''' </summary>
        ''' <remarks>
        ''' Listener must update their selection state as well as their lists, by 
        '''   querying ConfigurationDropdownEntries, PlatformDropdownEntries, 
        '''   SelectedConfigIndex and SelectedPlatformIndex.
        ''' </remarks>
        Private Sub ConfigurationState_ConfigurationListAndSelectionChanged()
            Debug.Assert(IsConfigPage)
            If IsConfigPage Then
                'Update out list
                UpdateConfigLists()

                '... and our selection state
                ChangeSelectedComboBoxIndicesWithoutNotification(_configurationState.SelectedConfigIndex, _configurationState.SelectedPlatformIndex)

                '.. and tell the page to update based on the new selection 'CONSIDER delaying this call until we're the active designer
                SetObjectsForSelectedConfigs()
            End If
        End Sub

        ''' <summary>
        ''' Raised when the undo/redo stack of a property page should be cleared because of
        '''   changes to configurations/platforms that are not currently supported by our
        '''   undo/redo story.
        ''' </summary>
        Private Sub ConfigurationState_ClearConfigPageUndoRedoStacks()
            Switches.TracePDUndo("Clearing undo/redo stack for page """ & GetPageTitle() & """")
            ClearUndoStackForPage()
        End Sub

        ''' <summary>
        ''' Raised when the value of the SimplifiedConfigMode property changes.
        ''' </summary>
        Private Sub ConfigurationState_SimplifiedConfigModeChanged()
            SetConfigDropdownVisibility()

            'This may change the objects selected.  Also, pages might have UI that depends on this setting, so get everything to update...
            SetObjectsForSelectedConfigs()
        End Sub

        ''' <summary>
        ''' Check if the simplified configs mode property has changed (we do this on WM_SETFOCUS, since there's no notification
        '''   of a change)
        ''' </summary>
        Private Sub CheckForModeChanges()
            If _configurationState IsNot Nothing AndAlso _fInitialized AndAlso _needToCheckForModeChanges Then
                _configurationState.CheckForModeChanges()
                _needToCheckForModeChanges = False
            End If
        End Sub

        ''' <summary>
        ''' Updates the configuration and platform combobox dropdown lists and selects the first entry in each list
        ''' </summary>
        Private Sub UpdateConfigLists()
            If Not IsConfigPage Then
                Exit Sub
            End If

            'Populate the dropdowns
            ConfigurationComboBox.Items.Clear()
            PlatformComboBox.Items.Clear()
            ConfigurationComboBox.BeginUpdate()
            PlatformComboBox.BeginUpdate()
            Try
                For Each ConfigEntry As ConfigurationState.DropdownItem In _configurationState.ConfigurationDropdownEntries
                    ConfigurationComboBox.Items.Add(ConfigEntry.DisplayName)
                Next
                For Each PlatformEntry As ConfigurationState.DropdownItem In _configurationState.PlatformDropdownEntries
                    PlatformComboBox.Items.Add(PlatformEntry.DisplayName)
                Next
            Finally
                ConfigurationComboBox.EndUpdate()
                PlatformComboBox.EndUpdate()
            End Try

            'Select the first entry in each combobox
            ChangeSelectedComboBoxIndicesWithoutNotification(0, 0)
        End Sub

        ''' <summary>
        ''' Returns the currently selected config combobox item
        ''' </summary>
        Private Function GetSelectedConfigItem() As ConfigurationState.DropdownItem
            Debug.Assert(ConfigurationComboBox.SelectedIndex >= 0)
            Debug.Assert(ConfigurationComboBox.Items.Count = _configurationState.ConfigurationDropdownEntries.Length,
                "The combobox is not in sync")
            Dim ConfigItem As ConfigurationState.DropdownItem = _configurationState.ConfigurationDropdownEntries(ConfigurationComboBox.SelectedIndex)
            Debug.Assert(ConfigItem IsNot Nothing)
            Return ConfigItem
        End Function

        ''' <summary>
        ''' Returns the currently selected platform combobox item
        ''' </summary>
        Private Function GetSelectedPlatformItem() As ConfigurationState.DropdownItem
            Debug.Assert(PlatformComboBox.SelectedIndex >= 0)
            Debug.Assert(PlatformComboBox.Items.Count = _configurationState.PlatformDropdownEntries.Length,
                "The combobox is not in sync")
            Dim PlatformItem As ConfigurationState.DropdownItem = _configurationState.PlatformDropdownEntries(PlatformComboBox.SelectedIndex)
            Debug.Assert(PlatformItem IsNot Nothing)
            Return PlatformItem
        End Function

        ''' <summary>
        ''' Determines the set of currently-selected configurations by inspecting the configuration comboboxes, 
        '''   and passes that set of objects to the loaded property page
        ''' </summary>
        Private Sub SetObjectsForSelectedConfigs()
            CommitPendingChanges()

            If Not IsConfigPage Then
                'Use the project's browse object for SetObjects
                CallPageSetObjects(New Object() {GetProjectBrowseObject()})
            Else
                'If here, then we are a config-dependent page...

                Dim SelectedConfigIndex As Integer = ConfigurationComboBox.SelectedIndex
                Debug.Assert(SelectedConfigIndex <> -1, "Selection should be set for config name")
                Dim SelectedConfigItem As ConfigurationState.DropdownItem = GetSelectedConfigItem()

                Dim SelectedPlatformIndex As Integer = PlatformComboBox.SelectedIndex
                Debug.Assert(SelectedPlatformIndex <> -1, "Selection should be set for platfofm")
                Dim SelectedPlatformItem As ConfigurationState.DropdownItem = GetSelectedPlatformItem()

                Dim AllConfigurations As Boolean = False
                If SelectedConfigItem.SelectionType = ConfigurationState.SelectionTypes.All Then
                    'User selected "All Configurations"
                    AllConfigurations = True
                ElseIf Not ConfigurationComboBox.Visible Then
                    ' When the config panel is hidden we should update all configs/platforms
                    AllConfigurations = True
                End If

                Dim AllPlatforms As Boolean = False
                If SelectedPlatformItem.SelectionType = ConfigurationState.SelectionTypes.All Then
                    'User selected "All Platforms"
                    AllPlatforms = True
                ElseIf Not ConfigurationComboBox.Visible Then
                    ' When the config panel is hidden we should update all configs/platforms
                    AllPlatforms = True
                End If

                'Find all matching config/platform combinations
                Dim ConfigObjects As Object() = Nothing

                If AllConfigurations AndAlso AllPlatforms Then
                    'All configurations and platforms
                    Dim Configs() As IVsCfg = _configurationState.GetAllConfigs()

                    'Must have an array of object, not IVsCfg
                    ConfigObjects = New Object(Configs.Length - 1) {}
                    Configs.CopyTo(ConfigObjects, 0)
                ElseIf Not AllConfigurations AndAlso Not AllPlatforms Then
                    'A single config/platform combination selected
                    Dim Cfg As IVsCfg = Nothing
                    If VSErrorHandler.Succeeded(VsCfgProvider.GetCfgOfName(SelectedConfigItem.Name, SelectedPlatformItem.Name, Cfg)) Then
                        ConfigObjects = New Object() {Cfg}
                    Else
                        ShowErrorMessage(My.Resources.Designer.GetString(My.Resources.Designer.PPG_ConfigNotFound_2Args, SelectedConfigItem.Name, SelectedPlatformItem.Name))
                    End If
                Else
                    'Use the DTE to find all the configs with a certain config name or platform name, then
                    '  look up the IVsCfg for those that were found
                    Dim DTEConfigs As EnvDTE.Configurations

                    If AllConfigurations Then
                        Debug.Assert(SelectedPlatformItem.SelectionType <> ConfigurationState.SelectionTypes.All)
                        DTEConfigs = DTEProject.ConfigurationManager.Platform(SelectedPlatformItem.Name)
                    Else
                        Debug.Assert(AllPlatforms)
                        Debug.Assert(SelectedConfigItem.SelectionType <> ConfigurationState.SelectionTypes.All)
                        DTEConfigs = DTEProject.ConfigurationManager.ConfigurationRow(SelectedConfigItem.Name)
                    End If
                    Debug.Assert(DTEConfigs IsNot Nothing AndAlso DTEConfigs.Count > 0)

                    Dim Cfg As IVsCfg = Nothing
                    ConfigObjects = New Object(DTEConfigs.Count - 1) {}
                    For i As Integer = 0 To DTEConfigs.Count - 1
                        Dim DTEConfig As EnvDTE.Configuration = DTEConfigs.Item(i + 1) '1-indexed
                        If VSErrorHandler.Succeeded(VsCfgProvider.GetCfgOfName(DTEConfig.ConfigurationName, DTEConfig.PlatformName, Cfg)) Then
                            ConfigObjects(i) = Cfg
                        Else
                            ShowErrorMessage(My.Resources.Designer.GetString(My.Resources.Designer.PPG_ConfigNotFound_2Args, SelectedConfigItem.Name, SelectedPlatformItem.Name))
                            ConfigObjects = Nothing
                            Exit For
                        End If
                    Next
                End If

                'Finally, call SetObjects with the selected configs

                If ConfigObjects Is Nothing OrElse ConfigObjects.Length = 0 Then
                    'There was an error collecting this info - unload the page
                    UnLoadPage()
                    Return
                End If

                CallPageSetObjects(ConfigObjects)
            End If
        End Sub

        ''' <summary>
        ''' Passes the given set of objects (configurations, etc.) to the property page via its IPropertyPage2.SetObjects method.
        '''   (For config-dependent pages, this is currently IVsCfg objects, for other pages a browse object,
        '''    but it could theoretically be just about anything with the project context necessary
        '''    for the page properties).
        ''' </summary>
        ''' <param name="Objects"></param>
        Private Sub CallPageSetObjects(Objects() As Object)
            Dim Count As UInteger = 0
            If Objects IsNot Nothing Then
                Debug.Assert(Objects.Length <= UInteger.MaxValue, "Whoa!  Muchos objects!")
                Debug.Assert(TypeOf Objects Is Object(), "Objects must be an array of Object, not an array of anything else!")
                Count = CUInt(Objects.Length)
            End If

            Debug.Assert(PropPage IsNot Nothing, "PropPage is Nothing")
            PropPage.SetObjects(Count, Objects)
        End Sub

#End Region

        ''' <summary>
        ''' Fired when the configuration combobox is dropped down.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub ConfigurationComboBox_DropDown(sender As Object, e As EventArgs) Handles ConfigurationComboBox.DropDown
            'Set the drop-down width to handle all the text entries in it
            SetComboBoxDropdownWidth(ConfigurationComboBox)
        End Sub

        ''' <summary>
        ''' Fired when the configuration combobox is dropped down.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub PlatformComboBox_DropDown(sender As Object, e As EventArgs) Handles PlatformComboBox.DropDown
            'Set the drop-down width to handle all the text entries in it
            SetComboBoxDropdownWidth(PlatformComboBox)
        End Sub

#Region "Undo/Redo handling"

        ''' <summary>
        ''' Gets the current value of the property on the specified PropPageDesignerRootComponent instance.
        ''' Used by the undo/redo mechanism to save current property values so that they can be 
        '''   automatically changed later via undo or redo.
        ''' </summary>
        ''' <param name="PropertyName">The name of the property on the current page to get the value of.</param>
        ''' <returns>The value of the property on the specified component instance.</returns>
        Public Function GetProperty(PropertyName As String) As Object
            If PropertyName = "" Then
                Throw New ArgumentException
            End If

            Switches.TracePDUndo("PropPageDesignerView.GetProperty(" & PropertyName & ")")

            Dim PropPageUndo As IVsProjectDesignerPage = TryCast(_loadedPage, IVsProjectDesignerPage)
            Debug.Assert(PropPageUndo IsNot Nothing)
            If PropPageUndo IsNot Nothing Then
                Dim Value As Object = Nothing

                'Only worry about multiple-value undo/redo for config-dependent pages...
                If IsConfigPage Then
                    If PropPageUndo.SupportsMultipleValueUndo(PropertyName) Then
                        Dim Objects As Object() = Nothing
                        Dim Values As Object() = Nothing

                        Switches.TracePDUndo("  Multi-value undo/redo supported by property.")
                        Try
                            'This page supports multiple-value undo/redo for this property (i.e., restoring different
                            '  values for each config), so attempt to get those values.
                            Dim GetValuesSucceeded As Boolean = PropPageUndo.GetPropertyMultipleValues(PropertyName, Objects, Values)
                            If GetValuesSucceeded AndAlso Objects IsNot Nothing AndAlso Values IsNot Nothing Then
                                'Package them up into a serializable class from which we can unpack them later during undo.
                                Dim SelectedConfigurationName As String = IIf(GetSelectedConfigItem().SelectionType = ConfigurationState.SelectionTypes.All, "", GetSelectedConfigItem().Name)
                                Dim SelectedPlatformName As String = IIf(GetSelectedPlatformItem().SelectionType = ConfigurationState.SelectionTypes.All, "", GetSelectedPlatformItem().Name)
                                Value = New MultipleValuesStore(VsCfgProvider, Objects, Values, SelectedConfigurationName, SelectedPlatformName)
                            Else
                                'GetPropertyMultipleValues returned Nothing.  Try for a single value later.
                            End If
                        Catch ex As NotSupportedException When ReportWithoutCrash(ex, "Prop page said it supported multiple value undo, but then failed with not supported", NameOf(PropPageDesignerView))
                            'Ignore error and try single value instead
                        Catch ex As ArgumentException
                            'Most likely this indicates that Objects were not IVsCfg (this could be the case for non-config-dependent pages).  We shouldn't 
                            '  have tried to call the multi-value undo stuff in this case, but if it does happen, let's tolerate 
                            '  it by reverting to single-value undo behavior.  MultipleValues will have already asserted in this case, so 
                            '  we don't need to unless this assumption is wrong.
                            Debug.Assert(Objects IsNot Nothing AndAlso Objects.Length = 1 AndAlso TypeOf Objects(0) IsNot IVsCfg,
                                "Unexpected exception in MultipleValues constructor.  Reverting to single-value undo/redo.")
                        End Try
                    Else
                        Switches.TracePDUndo("  Multi-value undo/redo not supported.")
                    End If
                Else
                    Switches.TracePDUndo("  Not a Config page, no multi-value undo.")
                    Debug.Assert(Not PropPageUndo.SupportsMultipleValueUndo(PropertyName),
                        "A property on a config-independent page supports multiple-value undo/redo.  That means the page contains a config-dependent property.  And that doesn't seem right, does it?" _
                        & vbCrLf & "PropertyName = " & PropertyName)
                End If

                'Getting multiple-value undo wasn't supported or didn't succeed.  Try getting just a single value.
                If Value Is Nothing Then
                    Value = PropPageUndo.GetProperty(PropertyName)
                End If

                'If any of the values being serialized is an unmanaged enum, the serialization stream will
                '  contain a reference to the dll, which the deserializer may have trouble deserializing.  So
                '  instead simply convert these to their underlying types (usually integer) to avoid complications.
                If TypeOf Value Is MultipleValuesStore Then
                    Dim Store As MultipleValuesStore = DirectCast(Value, MultipleValuesStore)
                    For i As Integer = 0 To Store.Values.Length - 1
                        ConvertEnumToUnderlyingType(Store.Values(i))
                    Next
                Else
                    ConvertEnumToUnderlyingType(Value)
                End If

                'Done.
                Switches.TracePDUndo("  Value=" & DebugToString(Value))
                Return Value
            Else
                Debug.Fail("PropertyPagePropertyDescriptor.GetValue() called with unexpected Component type.  Expected that this is also set up through the PropPageDesignerView (implementing IProjectDesignerPropertyPageUndoSite)")
                Throw New ArgumentException
            End If
        End Function

        ''' <summary>
        ''' Sets the value of a property on the active page to a different value.  Used during undo/redo.
        ''' </summary>
        ''' <param name="PropertyName">The property name of the property to be set</param>
        ''' <param name="Value">The value to set it to</param>
        ''' <remarks>
        ''' This method gets called by the serialization store dealing Undo/Redo operations.
        ''' </remarks>
        Public Sub SetProperty(PropertyName As String, Value As Object)
            If String.IsNullOrEmpty(PropertyName) Then
                Throw CreateArgumentException(NameOf(PropertyName))
            End If

            Switches.TracePDUndo("PropPageDesignerView.SetProperty(""" & PropertyName & """, " & DebugToString(Value) & ")")

            Dim PropPageUndo As IVsProjectDesignerPage = TryCast(_loadedPage, IVsProjectDesignerPage)
            Debug.Assert(PropPageUndo IsNot Nothing)
            If PropPageUndo IsNot Nothing Then
                'Is it a set of different values for multiple configurations?
                Dim MultiValues As MultipleValuesStore = TryCast(Value, MultipleValuesStore)
                If MultiValues IsNot Nothing Then
                    'Yes - multiple values need to be undone.
                    Switches.TracePDUndo("  Multi-value undo/redo.")

                    Debug.Assert(IsConfigPage, "How did we get multiple properties values for undo/redo for a non-config-dependent page?")
                    If Not PropPageUndo.SupportsMultipleValueUndo(PropertyName) Then
                        Debug.Fail("Property page that supported multi-value undo when saving the value doesn't support them now that we want to do the undo/redo?")
                    Else
                        'We are about to do an Undo or Redo.  Since the undo/redo data carries the configs/platforms that
                        '  were in effect when the original change was made, and we are about to revert to those changes,
                        '  we first need to select those same configurations again.
                        ReselectConfigurationsForUndoRedo(MultiValues)

                        'Tell the property page to set the new (or old) values
                        Dim Objects As Object() = MultiValues.GetObjects(VsCfgProvider)
                        Try
                            PropPageUndo.SetPropertyMultipleValues(PropertyName, Objects, MultiValues.Values)
                        Catch ex As NotSupportedException When ReportWithoutCrash(ex, "Property page threw not supported exception trying to undo/redo multi-value change", NameOf(PropPageDesignerView))
                        End Try
                    End If
                Else
                    'Nope - single value.  Since this is config-independent, no need to change the configuration/platform dropdowns
                    '  in this case.
                    PropPageUndo.SetProperty(PropertyName, Value)
                End If
            End If
        End Sub

        ''' <summary>
        ''' If the object passed in is an enum, then convert it to its underlying type
        ''' </summary>
        ''' <param name="Value">[inout] The value to check and convert in place.</param>
        Private Shared Sub ConvertEnumToUnderlyingType(ByRef Value As Object)
            If Value IsNot Nothing AndAlso Value.GetType().IsEnum Then
                Value = Convert.ChangeType(Value, Type.GetTypeCode(Value.GetType().UnderlyingSystemType))
            End If
        End Sub

        ''' <summary>
        ''' Returns the selection state of the configuration and platform comboboxes to the pre-undo/redo state.
        ''' </summary>
        ''' <param name="MultiValues">Specifies the selected configuration and platform in the drop-down combobox.</param>
        Private Sub ReselectConfigurationsForUndoRedo(MultiValues As MultipleValuesStore)
            If Not IsConfigPage Then
                Exit Sub
            End If

            Debug.Assert(MultiValues IsNot Nothing)

            Dim SelectAllConfigs As Boolean = MultiValues.SelectedConfigName = ""
            Dim SelectAllPlatforms As Boolean = MultiValues.SelectedPlatformName = ""

            _configurationState.ChangeSelection(
                MultiValues.SelectedConfigName, IIf(SelectAllConfigs, ConfigurationState.SelectionTypes.All, ConfigurationState.SelectionTypes.Normal),
                MultiValues.SelectedPlatformName, IIf(SelectAllPlatforms, ConfigurationState.SelectionTypes.All, ConfigurationState.SelectionTypes.Normal),
                PreferExactMatch:=False, FireNotifications:=True)
        End Sub

#End Region

        ''' <summary>
        ''' Commits any pending changes on the page
        ''' </summary>
        ''' <returns>return False if it failed</returns>
        Private Function CommitPendingChanges() As Boolean
            Switches.TracePDPerfBegin("PropPageDesignerView.CommitPendingChanges")
            Try
                If _loadedPageSite IsNot Nothing Then
                    If Not _loadedPageSite.CommitPendingChanges() Then
                        Return False
                    End If
                End If

                ' It is time to do all pending validations...
                Dim vbPropertyPage As IVsProjectDesignerPage = TryCast(_loadedPage, IVsProjectDesignerPage)
                If vbPropertyPage IsNot Nothing Then
                    If Not vbPropertyPage.FinishPendingValidations() Then
                        Return False
                    End If
                End If

                Return True
            Finally
                Switches.TracePDPerfEnd("PropPageDesignerView.CommitPendingChanges")
            End Try
        End Function

#Region "IVsWindowPaneCommit"
        ''' <summary>
        ''' This function is called on F5, build, etc., when any pending changes need to be performed on, say, a textbox
        '''   that the user has started typing into but hasn't committed by moving to another control.  We use this to force
        '''   an immediate apply.
        ''' </summary>
        ''' <param name="pfCommitFailed">[Out] Set to non-zero to indicate that the action should be canceled because the commit failed.</param>
        Public Function IVsWindowPaneCommit_CommitPendingEdit(ByRef pfCommitFailed As Integer) As Integer Implements IVsWindowPaneCommit.CommitPendingEdit
            pfCommitFailed = 0
            If Not CommitPendingChanges() Then
                pfCommitFailed = 1
            End If
            Return NativeMethods.S_OK
        End Function
#End Region

#Region "Message routing"

        ''' <summary>
        ''' Override this to enable tabbing from the property page designer (configuration panel) to
        '''   a native-hosted property page's controls.
        ''' </summary>
        ''' <param name="forward"></param>
        Protected Overrides Function ProcessTabKey(forward As Boolean) As Boolean
            Switches.TracePDMessageRouting(TraceLevel.Warning, "PropPageDesignerView.ProcessTabKey")

            If _isNativeHostedPropertyPage Then
                'Try tabbing to another control in the property page designer view
                If SelectNextControl(ActiveControl, forward, True, True, False) Then
                    Switches.TracePDMessageRouting(TraceLevel.Info, "  ...PropPageDesignerView.SelectNextControl handled it")
                    Return True
                End If

                If _loadedPage IsNot Nothing Then
                    'We hit the last tabbable control in the property page designer, set focus to the first (or last)
                    '  control in the property page itself.
                    Switches.TracePDMessageRouting(TraceLevel.Warning, "  ...Setting focus to " & IIf(forward, "first", "last") & " control on the page")
                    If Not FocusFirstOrLastPropertyPageControl(forward) Then
                        'No focusable controls in the property page (could be disabled), set focus to the
                        '  property page designer again
                        Return SelectNextControl(ActiveControl, forward, True, True, True)
                    End If
                    Return True
                End If
            Else
                If SelectNextControl(ActiveControl, forward, tabStopOnly:=True, nested:=True, wrap:=False) Then
                    Return True
                Else
                    Dim appDesView As ApplicationDesignerView = CType(_loadedPageSite.Owner, ApplicationDesignerView)
                    appDesView.SelectedItem.Focus()
                    appDesView.SelectedItem.FocusedFromKeyboardNav = True
                End If
            End If

            Return False
        End Function

        'For debug tracing
        Public Overrides Function PreProcessMessage(ByRef msg As Message) As Boolean
            Switches.TracePDMessageRouting(TraceLevel.Warning, "PropPageDesignerView.PreProcessMessage", msg)
            Return MyBase.PreProcessMessage(msg)
        End Function

        ''' <summary>
        ''' Override Control's ProcessDialogChar in order to allow mnemonics to work.
        ''' </summary>
        ''' <param name="charCode"></param>
        Protected Overrides Function ProcessDialogChar(charCode As Char) As Boolean
            'Control's version of this function only calls ProcessMnemonic if the window
            '  is top-level, but the control is not top-level as far as WinForms is concerned.
            '  We'll ensure that it's always called for us.
            If charCode <> " "c AndAlso ProcessMnemonic(charCode) Then
                Return True
            End If

            'If we're hosting a control in native, ProcessMnemonic will not have seen
            '  the native control, so we need to give the property page a crack at it
            If _isNativeHostedPropertyPage AndAlso _isPageActivated Then
                If charCode <> " "c AndAlso _loadedPage IsNot Nothing Then

                    'CONSIDER: theoretically we should allow non-alt accelerators, but only
                    '  if it's not an input key for the currently-active control, and there's
                    '  no way to get that info without late binding, so we'll just only
                    '  accept ALT accelerators.
                    If (ModifierKeys And Keys.Alt) <> 0 Then
                        Dim PropertyPageHwnd As IntPtr = GetPropertyPageTopHwnd()
                        Dim msg As OleInterop.MSG() = {New OleInterop.MSG}
                        With msg(0)
                            .hwnd = PropertyPageHwnd
                            .message = Win32Constant.WM_SYSCHAR
                            .wParam = New IntPtr(AscW(charCode))
                        End With
                        If _loadedPage.TranslateAccelerator(msg) = NativeMethods.S_OK Then
                            Return True
                        End If
                    End If
                End If
            End If

            Return MyBase.ProcessDialogChar(charCode)
        End Function

        ''' <summary>
        ''' Sets the focus to the first (or last) control in the property page.
        ''' </summary>
        ''' <param name="First"></param>
        Public Function FocusFirstOrLastPropertyPageControl(First As Boolean) As Boolean
            'Make sure to set the active control as well as doing SetFocus(), or else when
            '  devenv gets focus the focus will not go back to the correct control. 
            ActiveControl = PropertyPagePanel
            Return FocusFirstOrLastTabItem(GetPropertyPageTopHwnd(), First)
        End Function

#End Region

        ''' <summary>
        ''' Retrieves the top-most HWND of the property page hosted inside the property page panel
        ''' </summary>
        Public Function GetPropertyPageTopHwnd() As IntPtr
            If PropertyPagePanel.Handle.Equals(IntPtr.Zero) Then
                Return IntPtr.Zero
            End If

            Return NativeMethods.GetWindow(PropertyPagePanel.Handle, Win32Constant.GW_CHILD)
        End Function

        Public Sub OnActivated(activated As Boolean) Implements IVsEditWindowNotify.OnActivated
            Switches.TracePDPerfBegin("PropPageDesignerView.OnActivated")
            ' It is time to do all pending validations...
            Dim vbPropertyPage As IVsProjectDesignerPage = TryCast(_loadedPage, IVsProjectDesignerPage)
            If vbPropertyPage IsNot Nothing Then
                vbPropertyPage.OnActivated(activated)

                If activated Then
                    ' When an existing page is reactivated (i.e. switching back from something else),
                    ' reinitialize the ui cue state (This is like reopening a dialog where the state
                    ' is reinitialized)
                    InitializeStateOfUICues()
                End If
            End If
            If activated AndAlso _rootDesigner IsNot Nothing Then
                _rootDesigner.RefreshMenuStatus()
            End If

            Switches.TracePDPerfEnd("PropPageDesignerView.OnActivated")
        End Sub

        ''' <summary>
        ''' Clears all undo and redo entries for this page
        ''' </summary>
        Private Sub ClearUndoStackForPage()
            'Only need to do this if undo is enabled
            Dim UndoEngine As UndoEngine = TryCast(GetService(GetType(UndoEngine)), UndoEngine)
            Debug.Assert(UndoEngine IsNot Nothing, "Unable to get UndoEngine")
            If UndoEngine IsNot Nothing Then
                Debug.Assert(Not UndoEngine.UndoInProgress, "Trying to clear Undo stack while undo is in progress")
                If Not UndoEngine.UndoInProgress Then
                    Dim UndoManager As OleInterop.IOleUndoManager = TryCast(GetService(GetType(OleInterop.IOleUndoManager)), OleInterop.IOleUndoManager)
                    Debug.Assert(UndoManager IsNot Nothing, "Unable to get undo manager to clear the undo stack for the property page")
                    If UndoManager IsNot Nothing Then
                        Try
                            UndoManager.DiscardFrom(Nothing) 'Causes it to clear all entries in the undo and redo stacks
                        Catch ex As COMException When ReportWithoutCrash(ex, "Unable to clear the undo stack, perhaps a unit was open or in progress, or it is disabled?", NameOf(PropPageDesignerView))
                        End Try
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' Overridden WndProc
        ''' </summary>
        ''' <param name="m"></param>
        Protected Overrides Sub WndProc(ByRef m As Message)
            Dim isSetFocusMessage As Boolean = False
            If m.Msg = Win32Constant.WM_SETFOCUS Then
                isSetFocusMessage = True
            End If

            If isSetFocusMessage AndAlso _fInitialized Then
                ' NOTE: we stop auto-scroll to the active control when the whole view gets Focus, so we won't change the scrollbar position when the user switches between application and editors in the VS
                '  We should auto-scroll to the position when the page is just loaded.
                '  We should still scroll to the right view when focus is moving within the page.
                ' This is for vswhidbey: #517826
                PropertyPagePanel.StopAutoScrollToControl(True)
                Try
                    MyBase.WndProc(m)
                Finally
                    PropertyPagePanel.StopAutoScrollToControl(False)
                End Try
            Else
                MyBase.WndProc(m)
            End If

            If isSetFocusMessage Then
                'Since there's no notification of tools.option changes, on WM_SETFOCUS we check if the
                '  user has changed the simplified configs mode and update the page.
                If _configurationState IsNot Nothing AndAlso IsHandleCreated AndAlso Not _needToCheckForModeChanges Then
                    _needToCheckForModeChanges = True
                    BeginInvoke(New MethodInvoker(AddressOf CheckForModeChanges)) 'Make sure we're not in the middle of something when doing the check...
                End If
            End If
        End Sub

        ''' <summary>
        ''' Retrieves the title of the loaded property page
        ''' </summary>
        Private Function GetPageTitle() As String
            Dim Info As OleInterop.PROPPAGEINFO() = New OleInterop.PROPPAGEINFO(0) {}
            If _loadedPage IsNot Nothing Then
                _loadedPage.GetPageInfo(Info)
                Return Info(0).pszTitle
            End If

            Return ""
        End Function

        ''' <summary>
        ''' Sets the undo/redo level to a "clean" state
        ''' </summary>
        Public Sub SetUndoRedoCleanState()
            Dim CurrentUndoUnitsAvailable As Integer
            If TryGetUndoUnitsAvailable(CurrentUndoUnitsAvailable) Then
                'The page will be considered "clean" in the sense of undo/redo whenever
                '  the number of undo units available matches the number available right now.
                _undoUnitsOnStackAtCleanState = CurrentUndoUnitsAvailable
            Else
                Debug.Fail("SetUndoRedoCleanState(): unable to get undo units available")
                _undoUnitsOnStackAtCleanState = 0
            End If

            'For pages that don't support Undo/Redo, reset their flag
            _loadedPageSite.HasBeenSetDirty = False
        End Sub

        ''' <summary>
        ''' Determines whether the page is dirty in the sense of, "no current changes and the undo stack is at the same
        '''   place as when the user first loaded the page."
        ''' </summary>
        Public Function ShouldShowDirtyIndicator() As Boolean
            If PropPage Is Nothing Then
                ' This can happen if an exception happened during property page initialization
                Return False
            End If
            If _designerHost IsNot Nothing AndAlso _designerHost.InTransaction Then
                ' We will be called when the transaction closes...
                '
                Return False
            ElseIf TypeOf PropPage Is IVsProjectDesignerPage Then
                Dim UndoUnitsAvailable As Integer
                If TryGetUndoUnitsAvailable(UndoUnitsAvailable) Then
                    Return UndoUnitsAvailable <> _undoUnitsOnStackAtCleanState
                Else
                    'This can happen if the property page didn't load properly, etc.
                    Switches.TracePDUndo("*** ShouldShowDirtyIndicator: Returning FALSE because GetUndoUnitsAvailable failed (possibly couldn't get UndoEngine)")
                    Return False
                End If
            Else
                'Pages which do not support undo/redo simply show the asterisk if they are dirty
                If PropPage.IsPageDirty() = NativeMethods.S_OK Then
                    'Page is dirty
                    Return True
                End If

                ' ... or if the page has been marked dirty.  Some pages use immediate apply by sending us dirty+validate status,
                '  and those will get immediately set back to "clean".  This allows us to remember that they've actually been
                '  set as dirty by the user until the project designer is saved.
                If _loadedPageSite.HasBeenSetDirty Then
                    Return True
                End If

                Return False
            End If
        End Function

        Private Sub InitializeStateOfUICues()

            ' Passing UIS_INITIALIZE lets the OS decide what the initial state of the cue
            ' (whether they are hidden or not).  The cue flags are being bit shifted since
            ' WM_UPDATEUISTATE expects them in the hi order word of the wParam
            Dim updateUIStateWParam As Integer = NativeMethods.UIS_INITIALIZE Or
                                                 NativeMethods.UISF_HIDEFOCUS << 16 Or
                                                 NativeMethods.UISF_HIDEACCEL << 16

            NativeMethods.SendMessage(Handle, NativeMethods.WM_UPDATEUISTATE, New IntPtr(updateUIStateWParam), IntPtr.Zero)

        End Sub

#Region "Dummy disabled menu commands to let the proppages get keystrokes such as Ctrl+X"

        ''' <summary>
        ''' No-op command handler 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub DisabledMenuCommandHandler(sender As Object, e As EventArgs)
        End Sub
#End Region
        ''' <summary>
        ''' Returns the number of undo units currently available
        ''' </summary>
        ''' <returns>True if the function successfully retrieves the # of undo units available.  False if it
        '''   fails (e.g., the Undo engine is not available, which can happen when the property page didn't 
        '''   load properly, etc.)</returns>
        Private Function TryGetUndoUnitsAvailable(ByRef UndoUnitsAvailable As Integer) As Boolean
            UndoUnitsAvailable = 0

            Dim UndoEngine As UndoEngine = TryCast(GetService(GetType(UndoEngine)), UndoEngine)
            If UndoEngine IsNot Nothing Then
                Debug.Assert(Not UndoEngine.UndoInProgress, "Trying to get undo units while undo in progress")
                If Not UndoEngine.UndoInProgress Then
                    Dim UndoManager As OleInterop.IOleUndoManager = TryCast(GetService(GetType(OleInterop.IOleUndoManager)), OleInterop.IOleUndoManager)
                    Debug.Assert(UndoManager IsNot Nothing, "Unable to get IOleUndoManager from UneoEngine")
                    If UndoManager IsNot Nothing Then
                        Dim EnumUnits As OleInterop.IEnumOleUndoUnits = Nothing
                        UndoManager.EnumUndoable(EnumUnits)
                        If EnumUnits IsNot Nothing Then
                            Dim cUnits As Integer = 0
                            While True
                                Dim Units(0) As OleInterop.IOleUndoUnit
                                Dim cReturned As UInteger
                                If VSErrorHandler.Failed(EnumUnits.Next(1, Units, cReturned)) OrElse cReturned = 0 Then
                                    UndoUnitsAvailable = cUnits
                                    Return True
                                Else
                                    Debug.Assert(cReturned = 1)
                                    cUnits += 1
                                End If
                            End While
                        End If
                    End If
                End If
            End If

            Return False
        End Function

#If DEBUG Then
        Private Sub ConfigurationPanel_SizeChanged(sender As Object, e As EventArgs) Handles ConfigurationPanel.SizeChanged
            Switches.TracePDFocus(TraceLevel.Info, "ConfigurationPanel_SizeChanged: " & ConfigurationPanel.Size.ToString())
        End Sub
#End If

        ''' <summary>
        ''' Call when the property page panel gets focus
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub PropertyPagePanel_GotFocus(sender As Object, e As EventArgs) Handles PropertyPagePanel.GotFocus
            If _isNativeHostedPropertyPage Then
                'Since PropertyPagePanel has no child controls that WinForms knows about, we need to
                '  manually forward focus to the child.
                NativeMethods.SetFocus(NativeMethods.GetWindow(PropertyPagePanel.Handle, Win32Constant.GW_CHILD))
            End If
        End Sub

        ''' <summary>
        '''   We need disable scroll to the active control when the user switches application. 
        '''   We can do this by overriding ScrollToControl function, which is used to calculate preferred viewport position when one control is activated.
        ''' </summary>
        Public Class ScrollablePanel
            Inherits Panel

            ' whether we should disable auto-scroll the viewport to show the active control
            Private _stopAutoScrollToControl As Boolean

            ''' <summary>
            ''' change whether we should disable auto-scroll the viewport to show the active control
            ''' </summary>
            ''' <param name="needStop"></param>
            Public Sub StopAutoScrollToControl(needStop As Boolean)
                _stopAutoScrollToControl = needStop
            End Sub

            ''' <summary>
            ''' We overrides ScrollToControl to stop auto-scroll the viewport to show an active control when the customer switches between applications
            '''  The function is called to calculate the viewport to show the control. When we enable it, we let the base class to handle this correctly.
            '''  When we need disable the action, we simply return the current position of the view port, so the panel will not scroll automatically.
            ''' </summary>
            ''' <param name="activeControl"></param>
            Protected Overrides Function ScrollToControl(activeControl As Control) As Point
                If _stopAutoScrollToControl Then
                    Return DisplayRectangle.Location
                Else
                    Return MyBase.ScrollToControl(activeControl)
                End If
            End Function
        End Class
    End Class

End Namespace
