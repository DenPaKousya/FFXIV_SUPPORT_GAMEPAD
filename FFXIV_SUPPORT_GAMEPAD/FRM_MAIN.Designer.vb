<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FRM_MAIN
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FRM_MAIN))
        Me.NTI_TASK = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.CMS_MENU = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TSS1 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSM_SETTING = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSS2 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSM_EXIT = New System.Windows.Forms.ToolStripMenuItem()
        Me.TIM_CHECK_PROCESS = New System.Windows.Forms.Timer(Me.components)
        Me.CMS_MENU.SuspendLayout()
        Me.SuspendLayout()
        '
        'NTI_TASK
        '
        Me.NTI_TASK.ContextMenuStrip = Me.CMS_MENU
        Me.NTI_TASK.Visible = True
        '
        'CMS_MENU
        '
        Me.CMS_MENU.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CMS_MENU.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSS1, Me.TSM_SETTING, Me.TSS2, Me.TSM_EXIT})
        Me.CMS_MENU.Name = "CMS_MENU"
        Me.CMS_MENU.Size = New System.Drawing.Size(101, 60)
        '
        'TSS1
        '
        Me.TSS1.Name = "TSS1"
        Me.TSS1.Size = New System.Drawing.Size(97, 6)
        '
        'TSM_SETTING
        '
        Me.TSM_SETTING.Name = "TSM_SETTING"
        Me.TSM_SETTING.Size = New System.Drawing.Size(100, 22)
        Me.TSM_SETTING.Text = "設定"
        '
        'TSS2
        '
        Me.TSS2.Name = "TSS2"
        Me.TSS2.Size = New System.Drawing.Size(97, 6)
        '
        'TSM_EXIT
        '
        Me.TSM_EXIT.Name = "TSM_EXIT"
        Me.TSM_EXIT.Size = New System.Drawing.Size(100, 22)
        Me.TSM_EXIT.Text = "終了"
        '
        'TIM_CHECK_PROCESS
        '
        Me.TIM_CHECK_PROCESS.Enabled = True
        Me.TIM_CHECK_PROCESS.Interval = 500
        '
        'FRM_MAIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(10, 10)
        Me.ControlBox = False
        Me.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "FRM_MAIN"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.CMS_MENU.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents NTI_TASK As NotifyIcon
    Friend WithEvents CMS_MENU As ContextMenuStrip
    Friend WithEvents TSS1 As ToolStripSeparator
    Friend WithEvents TSM_SETTING As ToolStripMenuItem
    Friend WithEvents TSS2 As ToolStripSeparator
    Friend WithEvents TSM_EXIT As ToolStripMenuItem
    Friend WithEvents TIM_CHECK_PROCESS As Timer
End Class
