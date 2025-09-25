<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Fenster_Start
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Fenster_Start))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ListView = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.TabControl = New System.Windows.Forms.TabControl()
        Me.Tab1 = New System.Windows.Forms.TabPage()
        Me.PictureBox = New System.Windows.Forms.PictureBox()
        Me.Tab2 = New System.Windows.Forms.TabPage()
        Me.Panel_Info = New System.Windows.Forms.Panel()
        Me.Schaltflaeche_OK = New System.Windows.Forms.Button()
        Me.Text_Haftung = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label_Produkt_Version = New System.Windows.Forms.Label()
        Me.Label_Software_Name = New System.Windows.Forms.Label()
        Me.Panel_Info_Schatten = New System.Windows.Forms.Panel()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.Schaltflaeche_Beenden = New System.Windows.Forms.ToolStripMenuItem()
        Me.Schaltflaeche_PDFErzeugen = New System.Windows.Forms.ToolStripMenuItem()
        Me.Schaltflaeche_Info = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.TabControl.SuspendLayout()
        Me.Tab1.SuspendLayout()
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Info.SuspendLayout()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 300
        '
        'ListView
        '
        Me.ListView.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.ListView.HideSelection = False
        Me.ListView.LargeImageList = Me.ImageList
        Me.ListView.Location = New System.Drawing.Point(658, 62)
        Me.ListView.MultiSelect = False
        Me.ListView.Name = "ListView"
        Me.ListView.Size = New System.Drawing.Size(166, 891)
        Me.ListView.TabIndex = 12
        Me.ListView.UseCompatibleStateImageBehavior = False
        Me.ListView.Visible = False
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Wärmebrücke"
        '
        'ImageList
        '
        Me.ImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit
        Me.ImageList.ImageSize = New System.Drawing.Size(80, 120)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'TabControl
        '
        Me.TabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl.Controls.Add(Me.Tab1)
        Me.TabControl.Controls.Add(Me.Tab2)
        Me.TabControl.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl.Location = New System.Drawing.Point(12, 35)
        Me.TabControl.Name = "TabControl"
        Me.TabControl.SelectedIndex = 0
        Me.TabControl.Size = New System.Drawing.Size(638, 922)
        Me.TabControl.TabIndex = 14
        Me.TabControl.Visible = False
        '
        'Tab1
        '
        Me.Tab1.BackColor = System.Drawing.SystemColors.Control
        Me.Tab1.Controls.Add(Me.PictureBox)
        Me.Tab1.Location = New System.Drawing.Point(4, 27)
        Me.Tab1.Margin = New System.Windows.Forms.Padding(0)
        Me.Tab1.Name = "Tab1"
        Me.Tab1.Padding = New System.Windows.Forms.Padding(3)
        Me.Tab1.Size = New System.Drawing.Size(630, 891)
        Me.Tab1.TabIndex = 0
        Me.Tab1.Text = "Energieausweis"
        '
        'PictureBox
        '
        Me.PictureBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox.Name = "PictureBox"
        Me.PictureBox.Size = New System.Drawing.Size(624, 885)
        Me.PictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox.TabIndex = 13
        Me.PictureBox.TabStop = False
        Me.PictureBox.WaitOnLoad = True
        '
        'Tab2
        '
        Me.Tab2.BackColor = System.Drawing.SystemColors.Control
        Me.Tab2.Location = New System.Drawing.Point(4, 27)
        Me.Tab2.Name = "Tab2"
        Me.Tab2.Padding = New System.Windows.Forms.Padding(3)
        Me.Tab2.Size = New System.Drawing.Size(630, 891)
        Me.Tab2.TabIndex = 1
        Me.Tab2.Text = "Rückmeldung vom DIBt Server"
        '
        'Panel_Info
        '
        Me.Panel_Info.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Panel_Info.BackColor = System.Drawing.Color.White
        Me.Panel_Info.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Info.Controls.Add(Me.Schaltflaeche_OK)
        Me.Panel_Info.Controls.Add(Me.Text_Haftung)
        Me.Panel_Info.Controls.Add(Me.Label5)
        Me.Panel_Info.Controls.Add(Me.Label_Produkt_Version)
        Me.Panel_Info.Controls.Add(Me.Label_Software_Name)
        Me.Panel_Info.ForeColor = System.Drawing.Color.Black
        Me.Panel_Info.Location = New System.Drawing.Point(136, 251)
        Me.Panel_Info.Name = "Panel_Info"
        Me.Panel_Info.Size = New System.Drawing.Size(561, 479)
        Me.Panel_Info.TabIndex = 18
        Me.Panel_Info.Visible = False
        '
        'Schaltflaeche_OK
        '
        Me.Schaltflaeche_OK.Location = New System.Drawing.Point(469, 433)
        Me.Schaltflaeche_OK.Name = "Schaltflaeche_OK"
        Me.Schaltflaeche_OK.Size = New System.Drawing.Size(82, 37)
        Me.Schaltflaeche_OK.TabIndex = 6
        Me.Schaltflaeche_OK.Text = "OK"
        Me.Schaltflaeche_OK.UseVisualStyleBackColor = True
        '
        'Text_Haftung
        '
        Me.Text_Haftung.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text_Haftung.Location = New System.Drawing.Point(12, 118)
        Me.Text_Haftung.Multiline = True
        Me.Text_Haftung.Name = "Text_Haftung"
        Me.Text_Haftung.ReadOnly = True
        Me.Text_Haftung.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Text_Haftung.Size = New System.Drawing.Size(539, 309)
        Me.Text_Haftung.TabIndex = 5
        Me.Text_Haftung.Text = resources.GetString("Text_Haftung.Text")
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(18, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(442, 24)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "© Bundesamt für Bauwesen und Raumordnung"
        '
        'Label_Produkt_Version
        '
        Me.Label_Produkt_Version.AutoSize = True
        Me.Label_Produkt_Version.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Produkt_Version.Location = New System.Drawing.Point(18, 51)
        Me.Label_Produkt_Version.Name = "Label_Produkt_Version"
        Me.Label_Produkt_Version.Size = New System.Drawing.Size(218, 24)
        Me.Label_Produkt_Version.TabIndex = 1
        Me.Label_Produkt_Version.Text = "Produkt-Version: 1.0.1"
        '
        'Label_Software_Name
        '
        Me.Label_Software_Name.AutoSize = True
        Me.Label_Software_Name.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Software_Name.Location = New System.Drawing.Point(16, 13)
        Me.Label_Software_Name.Name = "Label_Software_Name"
        Me.Label_Software_Name.Size = New System.Drawing.Size(323, 34)
        Me.Label_Software_Name.TabIndex = 0
        Me.Label_Software_Name.Text = "Druckapplikation 2025"
        '
        'Panel_Info_Schatten
        '
        Me.Panel_Info_Schatten.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Panel_Info_Schatten.BackColor = System.Drawing.Color.LightGray
        Me.Panel_Info_Schatten.ForeColor = System.Drawing.Color.Black
        Me.Panel_Info_Schatten.Location = New System.Drawing.Point(146, 257)
        Me.Panel_Info_Schatten.Name = "Panel_Info_Schatten"
        Me.Panel_Info_Schatten.Size = New System.Drawing.Size(557, 479)
        Me.Panel_Info_Schatten.TabIndex = 19
        Me.Panel_Info_Schatten.Visible = False
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Schaltflaeche_Beenden, Me.Schaltflaeche_PDFErzeugen, Me.Schaltflaeche_Info})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(834, 29)
        Me.MenuStrip.TabIndex = 15
        Me.MenuStrip.Text = "MenuStrip1"
        '
        'Schaltflaeche_Beenden
        '
        Me.Schaltflaeche_Beenden.Name = "Schaltflaeche_Beenden"
        Me.Schaltflaeche_Beenden.Size = New System.Drawing.Size(82, 25)
        Me.Schaltflaeche_Beenden.Text = "Beenden"
        '
        'Schaltflaeche_PDFErzeugen
        '
        Me.Schaltflaeche_PDFErzeugen.Name = "Schaltflaeche_PDFErzeugen"
        Me.Schaltflaeche_PDFErzeugen.Size = New System.Drawing.Size(118, 25)
        Me.Schaltflaeche_PDFErzeugen.Text = "PDF erzeugen"
        '
        'Schaltflaeche_Info
        '
        Me.Schaltflaeche_Info.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.Schaltflaeche_Info.Name = "Schaltflaeche_Info"
        Me.Schaltflaeche_Info.Size = New System.Drawing.Size(49, 25)
        Me.Schaltflaeche_Info.Text = "Info"
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 959)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(834, 26)
        Me.StatusStrip.TabIndex = 16
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'StatusLabel
        '
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(119, 21)
        Me.StatusLabel.Text = "BBSR Drucktool"
        '
        'ProgressBar
        '
        Me.ProgressBar.ForeColor = System.Drawing.Color.ForestGreen
        Me.ProgressBar.Location = New System.Drawing.Point(366, 35)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(280, 24)
        Me.ProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar.TabIndex = 17
        Me.ProgressBar.Visible = False
        '
        'Fenster_Start
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(834, 985)
        Me.Controls.Add(Me.Panel_Info)
        Me.Controls.Add(Me.Panel_Info_Schatten)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.TabControl)
        Me.Controls.Add(Me.ListView)
        Me.Controls.Add(Me.MenuStrip)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "Fenster_Start"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BBSR Druckapplikation"
        Me.TabControl.ResumeLayout(False)
        Me.Tab1.ResumeLayout(False)
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Info.ResumeLayout(False)
        Me.Panel_Info.PerformLayout()
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As Windows.Forms.Timer
    Friend WithEvents ListView As Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As Windows.Forms.ColumnHeader
    Friend WithEvents ImageList As Windows.Forms.ImageList
    Friend WithEvents PictureBox As Windows.Forms.PictureBox
    Friend WithEvents TabControl As Windows.Forms.TabControl
    Friend WithEvents Tab1 As Windows.Forms.TabPage
    Friend WithEvents Tab2 As Windows.Forms.TabPage
    Friend WithEvents MenuStrip As Windows.Forms.MenuStrip
    Friend WithEvents Schaltflaeche_Beenden As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Schaltflaeche_PDFErzeugen As Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip As Windows.Forms.StatusStrip
    Friend WithEvents StatusLabel As Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ProgressBar As Windows.Forms.ProgressBar
    Friend WithEvents Schaltflaeche_Info As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel_Info As Windows.Forms.Panel
    Friend WithEvents Label_Produkt_Version As Windows.Forms.Label
    Friend WithEvents Label_Software_Name As Windows.Forms.Label
    Friend WithEvents Text_Haftung As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Schaltflaeche_OK As Windows.Forms.Button
    Friend WithEvents Panel_Info_Schatten As Windows.Forms.Panel
End Class
