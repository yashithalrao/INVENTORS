AddReference "System.Windows.Forms"
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports DrawingColor = System.Drawing.Color
Imports DrawingPoint = System.Drawing.Point
Imports WinFormsApp = System.Windows.Forms.Application
Imports System.Drawing.Drawing2D
Sub Main()
	
	' Get the active Inventor document
Dim invDoc = ThisApplication.ActiveDocument

' Ensure the rule runs only in a 3D model (Part or Assembly)
If Not (invDoc.DocumentType = 12291 Or invDoc.DocumentType = 12290) Then
    Return ' Exit if it's not a Part or Assembly
End If
' Add required references
	
    ' Create Form
    Dim form As New Form()
    form.Text = "i~ADAS"
     form.BackColor = DrawingColor.White  ' LightBlue
    form.Size = New Size(1400, 870)
    form.StartPosition = FormStartPosition.CenterScreen
     form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False



    '[ Set Background Image
    Dim docPath1 As String = System.IO.Path.GetDirectoryName(ThisApplication.ActiveDocument.FullFileName)
    Dim backgroundImagePath As String = System.IO.Path.Combine(docPath1, "IMAGES", "Untitled design.png")

    If System.IO.File.Exists(backgroundImagePath) Then
        form.BackgroundImage = Image.FromFile(backgroundImagePath)
        form.BackgroundImageLayout = ImageLayout.Stretch
    Else
        MessageBox.Show("Background image not found at: " & backgroundImagePath, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End If


    ' Create PictureBox for Forging Model
     Dim imgFM As New PictureBox()
     Dim imagePathFM As String = System.IO.Path.Combine(docPath1, "IMAGES", "component f.png")

    If System.IO.File.Exists(imagePathFM) Then
        imgFM.Image = Image.FromFile(imagePathFM)
    Else
        MessageBox.Show("Image file not found at: " & imagePathFM, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End If
    imgFM.Size = New Size(280, 230)
    imgFM.SizeMode = PictureBoxSizeMode.StretchImage
     imgFM.Location = New DrawingPoint(20, 50)
imgFM.BackColor = System.Drawing.Color.Transparent

 
    ' Create PictureBox for Forging Model
    Dim imgSAN As New PictureBox()
    Dim imagePathSAN As String = System.IO.Path.Combine(docPath1, "LOGO", "SanseraLogo (1).png")

    If System.IO.File.Exists(imagePathSAN) Then
         imgSAN.Image = Image.FromFile(imagePathSAN)
    Else
        MessageBox.Show("Image file not found at: " & imagePathSAN, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End If
    imgSAN.Size = New Size(270, 45)
    imgSAN.SizeMode = PictureBoxSizeMode.StretchImage
    imgSAN.Location = New DrawingPoint(1080, 30)
    imgSAN.BackColor = System.Drawing.Color.Transparent

 
    ' Create PictureBox for Machine Model
    Dim imgMM As New PictureBox()
    Dim imagePathMM As String = System.IO.Path.Combine(docPath1, "IMAGES", "component m.png")

     If System.IO.File.Exists(imagePathMM) Then
         imgMM.Image = Image.FromFile(imagePathMM)
    Else
        MessageBox.Show("Image file not found at: " & imagePathMM, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
     End If
    imgMM.Size = New Size(280, 230)
    imgMM.SizeMode = PictureBoxSizeMode.StretchImage
    imgMM.Location = New DrawingPoint(1050, 595)
	 imgMM.BackColor = System.Drawing.Color.Transparent
	 
	'] 
	 
   '[ IMAGE BUTTON	  
'    Dim imgB1 As New PictureBox()
'    Dim imagePathB1 As String = System.IO.Path.Combine(docPath1, "LOGO", "button 1 -dimension extraction.png")
  
'    If System.IO.File.Exists(imagePathB1) Then
'        imgB1.Image = Image.FromFile(imagePathB1)
'    Else
'        MessageBox.Show("Image file not found at: " & imagePathB1, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'     End If
'    imgB1.Size = New Size(250, 67.5)
'    imgB1.SizeMode = PictureBoxSizeMode.StretchImage
'    imgB1.Location = New DrawingPoint(630, 270)
'	 imgB1.BackColor = System.Drawing.Color.Transparent
'	 imgB1.Cursor = Cursors.Hand
'  	imgB1.Tag = "B1"
	

'AddHandler imgB1.MouseEnter, AddressOf imgB1_MouseEnter
' AddHandler imgB1.MouseLeave, AddressOf imgB1_MouseLeave

'' **Fix: Add Click event to close the form**
'AddHandler imgB1.Click, Sub(sender As Object, E As EventArgs)
'                            ' Close the parent form when the PictureBox is clicked
'                            Dim parentForm As form = imgB1.FindForm()
'                            If parentForm IsNot Nothing Then parentForm.Close()
'                        End Sub


Dim imgCD As New PictureBox()
 Dim imagePathCD As String = System.IO.Path.Combine(docPath1, "LOGO", "button- component details.png")

If System.IO.File.Exists(imagePathCD) Then
    imgCD.Image = Image.FromFile(imagePathCD)
Else
    MessageBox.Show("Image file not found at: " & imagePathCD, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
End If

imgCD.Size = New Size(250, 67.5)
imgCD.SizeMode = PictureBoxSizeMode.StretchImage
imgCD.Location = New DrawingPoint(700, 350)
imgCD.BackColor = System.Drawing.Color.Transparent
imgCD.Cursor = Cursors.Hand
imgCD.Tag = "CD"  ' Changed from "B1" to "CD"

' Assign event handlers for hover effect
AddHandler imgCD.MouseEnter, AddressOf imgCD_MouseEnter
AddHandler imgCD.MouseLeave, AddressOf imgCD_MouseLeave


Dim imgIFM As New PictureBox()
Dim imagePathIFM As String = System.IO.Path.Combine(docPath1, "LOGO", "import forging model.png")

If System.IO.File.Exists(imagePathIFM) Then
    imgIFM.Image = Image.FromFile(imagePathIFM)
Else
    MessageBox.Show("Image file not found at: " & imagePathIFM, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
End If

imgIFM.Size = New Size(250, 67.5)
imgIFM.SizeMode = PictureBoxSizeMode.StretchImage
imgIFM.Location = New DrawingPoint(800, 430) ' Fixed DrawingPoint to Point
imgIFM.BackColor = System.Drawing.Color.Transparent
  imgIFM.Cursor = Cursors.Hand
imgIFM.Tag = "IFM"  ' Changed from "CD" to "IFM"

' Assign event handlers for hover effect
AddHandler imgIFM.MouseEnter, AddressOf imgIFM_MouseEnter
AddHandler imgIFM.MouseLeave, AddressOf imgIFM_MouseLeave



Dim imgM2 As New PictureBox()
Dim imagePathM2 As String = System.IO.Path.Combine(docPath1, "LOGO", "generator.png")

If System.IO.File.Exists(imagePathFM) Then
    imgM2.Image = Image.FromFile(imagePathM2)
Else
    MessageBox.Show("Image file not found at: " & imagePathM2, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
End If

imgM2.Size = New Size(280, 67.5)
imgM2.SizeMode = PictureBoxSizeMode.StretchImage
imgM2.Location = New DrawingPoint(610, 600) ' Fixed DrawingPoint to Point
imgM2.BackColor = System.Drawing.Color.Transparent
imgM2.Cursor = Cursors.Hand
imgM2.Tag = "M2"  ' Changed from "CD" to "IFM"

' Assign event handlers for hover effect
AddHandler imgM2.MouseEnter, AddressOf imgM2_MouseEnter
AddHandler imgM2.MouseLeave, AddressOf imgM2_MouseLeave


'Dim imgIA As New PictureBox()
'Dim imagePathIA As String = System.IO.Path.Combine(docPath1, "LOGO", "ia assistance.png")

'If System.IO.File.Exists(imagePathIA) Then
'    imgIA.Image = Image.FromFile(imagePathIA)
'Else
'    MessageBox.Show("Image file not found at: " & imagePathIA, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'End If

'imgIA.Size = New Size(150, 150)
'imgIA.SizeMode = PictureBoxSizeMode.StretchImage
'imgIA.Location = New DrawingPoint(880, 650) ' Fixed DrawingPoint to Point
'imgIA.BackColor = System.Drawing.Color.Transparent
'imgIA.Cursor = Cursors.Hand
'imgIA.Tag = "IA"  ' Changed from "CD" to "IFM"

'' Assign event handlers for hover effect
'AddHandler imgIA.MouseEnter, AddressOf imgIA_MouseEnter
'AddHandler imgIA.MouseLeave, AddressOf imgIA_MouseLeave

']

    '[ Create Labels
     Dim label As New Label()
     label.Text = "i~ADAS"
    label.AutoSize = True
    label.Location = New DrawingPoint(600, 50)
    label.Font = New Font("Times New Roman", 30, FontStyle.Bold)
	label.BackColor = System.Drawing.Color.Transparent

    Dim label1 As New Label()
    label1.Text = "INTEGRAL CONNECTING ROD"
     label1.AutoSize = True
     label1.Location = New DrawingPoint(500, 120)
    label1.Font = New Font("Times New Roman", 22, FontStyle.Regular)
label1.BackColor = System.Drawing.Color.Transparent

 Dim labelM1 As New Label()
    labelM1.Text = "Module-1"
     labelM1.AutoSize = True
     labelM1.Location = New DrawingPoint(650, 220)
     labelM1.Font = New Font("Times New Roman", 20, FontStyle.Bold)
labelM1.BackColor = System.Drawing.Color.Transparent

 Dim labelM2 As New Label()
    labelM2.Text = "Module-2"
     labelM2.AutoSize = True
     labelM2.Location = New DrawingPoint(650, 550)
     labelM2.Font = New Font("Times New Roman", 20, FontStyle.Bold)
labelM2.BackColor = System.Drawing.Color.Transparent

 Dim labelIA As New Label()
    labelIA.Text = "Click here for IA Assistance"
     labelIA.AutoSize = True
     labelIA.Location = New DrawingPoint(500, 700)
     labelIA.Font = New Font("Times New Roman", 20, FontStyle.Bold)
labelIA.BackColor = System.Drawing.Color.Transparent


']

    '[ Create Buttons

Dim btnYes As New Button()
btnYes.Text = "" ' Remove text if using an image-only button
btnYes.Size = New Size(150, 150) ' Set a square size for a perfect circle
btnYes.Location = New DrawingPoint(880, 650)

' Make the button round
Dim path As New GraphicsPath()
path.AddEllipse(0, 0, btnYes.Width, btnYes.Height)
btnYes.Region = New Region(path)

' Load image from file
Dim imagePath As String = System.IO.Path.Combine(docPath1, "LOGO", "ia assistance.png")
If System.IO.File.Exists(imagePath) Then
    btnYes.BackgroundImage = Image.FromFile(imagePath)
    btnYes.BackgroundImageLayout = ImageLayout.Stretch ' Adjust image to fit button size
Else
    MessageBox.Show("Image not found: " & imagePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
End If

btnYes.FlatStyle = FlatStyle.Flat ' Optional: Remove border for better appearance
btnYes.FlatAppearance.BorderSize = 0

btnYes.DialogResult = DialogResult.Yes



Dim btnContinue As New Button()
btnContinue.Text = "" ' Remove text if using an image-only button
btnContinue.Size = New Size(250, 67.5) ' Wider slot shape
btnContinue.Location = New Drawing.Point(630 , 270) ' Button position
	 btnContinue.BackColor = System.Drawing.Color.Transparent
	 
' Create a slot shape (rounded rectangle)
Dim path10 As New Drawing.Drawing2D.GraphicsPath()
Dim rect As New Rectangle(0, 0, btnContinue.Width, btnContinue.Height)
Dim cornerRadius As Integer = 33.75 ' Adjust for more/less rounding

' Define slot shape with straight sides and rounded corners
path10.AddArc(rect.X, rect.Y, cornerRadius, cornerRadius, 180, 90) ' Top-left
path10.AddLine(rect.X + cornerRadius, rect.Y, rect.Right - cornerRadius, rect.Y) ' Top edge
path10.AddArc(rect.Right - cornerRadius, rect.Y, cornerRadius, cornerRadius, 270, 90) ' Top-right
path10.AddLine(rect.Right, rect.Y + cornerRadius, rect.Right, rect.Bottom - cornerRadius) ' Right edge
path10.AddArc(rect.Right - cornerRadius, rect.Bottom - cornerRadius, cornerRadius, cornerRadius, 0, 90) ' Bottom-right
path10.AddLine(rect.Right - cornerRadius, rect.Bottom, rect.X + cornerRadius, rect.Bottom) ' Bottom edge
path10.AddArc(rect.X, rect.Bottom - cornerRadius, cornerRadius, cornerRadius, 90, 90) ' Bottom-left
path10.AddLine(rect.X, rect.Bottom - cornerRadius, rect.X, rect.Y + cornerRadius) ' Left edge
path10.CloseFigure()

' Apply slot shape
btnContinue.Region = New Region(path10)
btnContinue.FlatStyle = FlatStyle.Flat ' Optional: Remove border for better appearance
btnContinue.FlatAppearance.BorderSize = 0

btnContinue.DialogResult = DialogResult.Ignore

' Load image from file
Dim imagePath10 As String = System.IO.Path.Combine(docPath1, "LOGO", "button 1 -dimension extraction.png")
If System.IO.File.Exists(imagePath10) Then
    btnContinue.BackgroundImage = Image.FromFile(imagePath10)
    btnContinue.BackgroundImageLayout = ImageLayout.Stretch

End If
 
Dim btnOk As New Button()
btnOk.Text = "" ' Remove text if using an image-only button
btnOk.Size = New Size(280, 67.5) ' Wider slot shape
btnOk.Location = New Drawing.Point(610, 600) ' Button position

' Create a slot shape (rounded rectangle)
Dim path11 As New Drawing.Drawing2D.GraphicsPath()
Dim rect11 As New Rectangle(0, 0, btnOk.Width, btnOk.Height)
Dim cornerRadius11 As Integer = 33.75 ' Adjust for more/less rounding

' Define slot shape with straight sides and rounded corners
path11.AddArc(rect11.X, rect11.Y, cornerRadius11, cornerRadius11, 180, 90) ' Top-left
path11.AddLine(rect11.X + cornerRadius11, rect11.Y, rect11.Right - cornerRadius11, rect11.Y) ' Top edge
path11.AddArc(rect11.Right - cornerRadius11, rect11.Y, cornerRadius11, cornerRadius11, 270, 90) ' Top-right
path11.AddLine(rect11.Right, rect11.Y + cornerRadius11, rect11.Right, rect11.Bottom - cornerRadius11) ' Right edge
path11.AddArc(rect11.Right - cornerRadius11, rect11.Bottom - cornerRadius11, cornerRadius11, cornerRadius11, 0, 90) ' Bottom-right
path11.AddLine(rect11.Right - cornerRadius11, rect11.Bottom, rect11.X + cornerRadius11, rect11.Bottom) ' Bottom edge
path11.AddArc(rect11.X, rect11.Bottom - cornerRadius11, cornerRadius11, cornerRadius11, 90, 90) ' Bottom-left
path11.AddLine(rect11.X, rect11.Bottom - cornerRadius11, rect11.X, rect11.Y + cornerRadius11) ' Left edge
path11.CloseFigure()



btnOk.Region = New Region(path11)
btnOk.FlatStyle = FlatStyle.Flat ' Optional: Remove border for better appearance
btnOk.FlatAppearance.BorderSize = 0

btnOk.DialogResult = DialogResult.Yes

 'Load image From file
Dim imagePath11 As String = System.IO.Path.Combine(docPath1, "LOGO", "generator.png") ' Update to a cancel button image if available
If System.IO.File.Exists(imagePath11) Then
    btnOk.BackgroundImage = Image.FromFile(imagePath11)
    btnOk.BackgroundImageLayout = ImageLayout.Stretch
End If

' Button Appearance
btnOk.FlatStyle = FlatStyle.Flat
btnOk.FlatAppearance.BorderSize = 0
btnOk.BackColor = System.Drawing.Color.Transparent
btnOk.DialogResult = DialogResult.OK ' Set to Cancel instead of OK
 


    Dim btnPD As New Button()
    btnPD.Text = "Fill Parts Details"
    btnPD.Location = New DrawingPoint(500, 370)
    btnPD.AutoSize = True
    btnPD.Font = New Font("Times New Roman", 12, FontStyle.Bold)
    btnPD.Tag = "PD"
 
    Dim btnIFM As New Button()
    btnIFM.Text = "Import Forging Model"
    btnIFM.Location = New DrawingPoint(500, 450)
    btnIFM.AutoSize = True
    btnIFM.Font = New Font("Times New Roman", 12, FontStyle.Bold)
    btnIFM.Tag = "IFM"

    Dim btnIFF As New Button()
    btnIFF.Text = "PFD/PD/CP/Fixture Design/CTE Generator"
    btnIFF.Location = New DrawingPoint(460, 540)
    btnIFF.AutoSize = True
    btnIFF.Font = New Font("Times New Roman", 12, FontStyle.Bold)
      btnIFF.Tag = "FF"

    ' Add Click Event Handlers
   'AddHandler btnDE.Click, AddressOf Button_Click
    AddHandler btnPD.Click, AddressOf Button_Click
     AddHandler btnIFM.Click, AddressOf Button_Click
    'AddHandler btnIFF.Click, AddressOf Button_Click
	'AddHandler btnYes.Click, AddressOf Button_Click
']

  ' Assign click event handler for imgB1
'AddHandler imgB1.Click, AddressOf imgB1_Click 
AddHandler imgIFM.Click, AddressOf imgIFM_Click
AddHandler imgCD.Click, AddressOf imgCD_Click
'AddHandler imgM2.Click, AddressOf imgM2_Click
'AddHandler imgIA.Click, AddressOf imgIA_Click


    ' Add Controls to the Form
    'form.Controls.Add(imgiadas)
	form.Controls.Add(imgSAN)
    form.Controls.Add(imgFM)
    form.Controls.Add(imgMM)
	'form.Controls.Add(imgB1)
	 form.Controls.Add(imgCD)
	 	 form.Controls.Add(imgIFM)
	
		 form.Controls.Add(imgIA)
    form.Controls.Add(label)
    form.Controls.Add(label1)
	 form.Controls.Add(labelM1)
	 form.Controls.Add(labelM2)
	 form.Controls.Add(labelIA)
    'form.Controls.Add(btnDE)
    form.Controls.Add(btnYes)
    form.Controls.Add(btnContinue)
    form.Controls.Add(btnOk)

    ' Show the Form
    
	
	  ' Set default buttons
form.AcceptButton = btnYes
'form.Accept2Button = btnYes2
'form.Button = btnContinue
'form.Button = btnOk
' Show the Form & Capture Result
Dim result As DialogResult = form.ShowDialog()

'			'voiceIA
'Dim objVoiceIA As Object = CreateObject("SAPI.SpVoice")
'objVoiceIA.Speak("Do You Wish To Continue With IA Voice Assistant", 1)
'' Custom Voice Toggle Form with Yes/No Buttons
'Dim voiceForm As New form
'voiceForm.Text = "i~ADAS Voice Control"
'voiceForm.BackColor = DrawingColor.LightBlue
'voiceForm.Width = 300
'voiceForm.Height = 150
'voiceForm.StartPosition = FormStartPosition.CenterScreen

'' Create a label for instructions
'Dim lblPrompt As New label
'lblPrompt.Text = "Do You Wish To Continue With IA Voice Assistant?"
'lblPrompt.AutoSize = True
'lblPrompt.Left = 10
'lblPrompt.Top = 30
'voiceForm.Controls.Add(lblPrompt)
 
'' Create Yes button
'Dim btnVoiceYes As New Button
'btnVoiceYes.Text = "Yes"
'btnVoiceYes.Left = 50
'btnVoiceYes.Top = 70
'AddHandler btnVoiceYes.Click, Sub(sender, E)
'    VoiceEnabled = True
'    voiceForm.Close()
'End Sub
'voiceForm.Controls.Add(btnVoiceYes)

'' Create No button
'Dim btnVoiceNo As New Button
'btnVoiceNo.Text = "No"
'btnVoiceNo.Left = 150
'btnVoiceNo.Top = 70
'AddHandler btnVoiceNo.Click, Sub(sender, E)
'    VoiceEnabled = False
	
'	Dim voiceOff As New form
'voiceOff.Text = "i~ADAS Voice Control Turned Off"
'voiceOff.BackColor = DrawingColor.LightBlue
'voiceOff.Width = 300
'voiceOff.Height = 150
'voiceOff.StartPosition = FormStartPosition.CenterScreen

'Dim OffPrompt As New label
'OffPrompt.Text = "i~ADAS Voice Control Turned Off"
'OffPrompt.AutoSize = True
'OffPrompt.Left = 10
'OffPrompt.Top = 30
'voiceOff.Controls.Add(OffPrompt)

'Dim btnOffOk As New Button
'btnOffOk.Text = "OK"
'btnOffOk.Left = 50
'btnOffOk.Top = 70
'voiceOff.Controls.Add(btnOffOk)
'AddHandler btnOffOk.Click, Sub(s, ev) voiceOff.Close()

'voiceOff.ShowDialog()
	
'    voiceForm.Close()
'End Sub
'voiceForm.Controls.Add(btnVoiceNo)

'' Show the form
'voiceForm.ShowDialog()

' Show the form and store the button click result
'Dim result As DialogResult = Me.ShowDialog()

' Handle Button Click Results
Select Case result
    Case DialogResult.Yes
        iLogicVb.RunRule("Main form")

    Case DialogResult.OK
        iLogicVb.RunRule("SHOW FORM(OVER ALL PARAMETERS 1)")
  Case DialogResult.Ignore
iLogicVb.RunRule("OPEN 3D MODEL(STEP-1)_Manual")

    Case Else
        ' Uncomment below if you want to show another form when no button is clicked
        ' form4.ShowDialog()
End Select

End Sub


	

	
'[OPEN 3D MODEL(STEP-1)
' MouseEnter: Highlight when hovered
Private Sub imgB1_MouseEnter(sender As Object, e As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.FixedSingle  ' Add a border to highlight
End Sub

' MouseLeave: Remove highlight when the cursor leaves
Private Sub imgB1_MouseLeave(sender As Object, e As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.None  ' Remove the border
End Sub']


Private Sub imgB1_Click(sender As Object, E As EventArgs)
    iLogicVb.RunRule("OPEN 3D MODEL(STEP-1)")
End Sub




'[DRAWING PARAMETERS
' MouseEnter: Highlight when hovered
Private Sub imgCD_MouseEnter(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.FixedSingle  ' Add a border to highlight
End Sub

' MouseLeave: Remove highlight when the cursor leaves
Private Sub imgCD_MouseLeave(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.None  ' Remove the border
End Sub
']

Private Sub imgCD_Click(sender As Object, E As EventArgs)
   iLogicVb.RunRule("SHOW FORM(DRAWING PARAMETERS)")

	
End Sub


'[ FORGING MODEL REPLACE
Private Sub imgIFM_MouseLeave(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.None  ' Remove the border
End Sub
 
' Click event: Show "DRAWING PARAMETERS" in iLogicForm
Private Sub imgIFM_Click(sender As Object, e As EventArgs)
   iLogicVb.RunRule("FORGING MODEL REPLACE")
End Sub
']

Private Sub imgIFM_MouseEnter(sender As Object, E As EventArgs)
     Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.FixedSingle  ' Add a border to highlight
End Sub

'[ MODULE-2
Private Sub imgM2_MouseLeave(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.None  ' Remove the border
End Sub

' Click event: Show "DRAWING PARAMETERS" in iLogicForm
Private Sub imgM2_Click(sender As Object, E As EventArgs)
   iLogicVb.RunRule("SHOW FORM(OVER ALL PARAMETERS 1)")
End Sub
']

Private Sub imgM2_MouseEnter(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.FixedSingle  ' Add a border to highlight
End Sub

'[OPEN 3D MODEL(STEP-1)
' MouseEnter: Highlight when hovered
Private Sub imgIA_MouseEnter(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.FixedSingle  ' Add a border to highlight
End Sub

' MouseLeave: Remove highlight when the cursor leaves
Private Sub imgIA_MouseLeave(sender As Object, E As EventArgs)
    Dim pb As PictureBox = CType(sender, PictureBox)
    pb.BorderStyle = BorderStyle.None  ' Remove the border
End Sub']

Private Sub imgIA_Click(sender As Object, E As EventArgs)
    iLogicVb.RunRule("Main form")
End Sub



Private Sub Button_Click(sender As Object, e As EventArgs)
    Dim btn As Button = CType(sender, Button)

    Select Case btn.Tag.ToString().ToUpper()
	
        Case "DE"
            iLogicVb.RunRule("OPEN 3D MODEL(STEP-1)")

        Case "PD"
            iLogicForm.Show("DRAWING PARAMETERS")

        Case "IFM"
            iLogicVb.RunRule("FORGING MODEL REPLACE")

        Case "FF"
            iLogicVb.RunRule("SHOW FORM(OVER ALL PARAMETERS 1)")




        Case Else
            Dim form4 As New Form()
            form4.Text = "i~ADAS"
            form4.BackColor = DrawingColor.FromArgb(255, 255, 102, 102)
            form4.Size = New Size(400, 120)
            form4.StartPosition = FormStartPosition.CenterScreen
            form4.FormBorderStyle = FormBorderStyle.FixedDialog

            Dim label5 As New Label()
            label5.Text = "Oops! Unknown command."
            label5.AutoSize = True
            label5.Location = New DrawingPoint(30, 25)
            label5.Font = New Font("Arial", 10, FontStyle.Bold)

            form4.Controls.Add(label5)
            form4.ShowDialog()
    End Select


' Get the active Inventor document
Dim invDoc = ThisApplication.ActiveDocument

' Ensure the rule runs only in a 3D model (Part or Assembly)
If Not (invDoc.DocumentType = 12291 Or invDoc.DocumentType = 12290) Then
    Return ' Exit if it's not a Part or Assembly
End If


iLogicVb.UpdateWhenDone = True
'Dim myForm As New Form()
'myForm.Close() ' Close the specific form instance
End Sub
