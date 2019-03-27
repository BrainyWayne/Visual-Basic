
'LAB4 WITH FUNCTIONS

Public Class Form1

    Private Cost As Double
    Private Salestax As Double
    Private Shippingfee As Double
    Private Handlingfee As Double
    Private valid = True
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtName.Focus()
        chkboxWholesale.Checked = "False"
        txtstate.CharacterCasing = CharacterCasing.Upper
        rduspostal.Checked = True
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles Customer_Information.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Computebtn.Click
        Dim total As Double
        ValidData()
        Cost = (txtPPP.Text * txtquantity.Text).ToString("n1")
        ComputeSalesTaxDue()
        ComputeTransportHandling()

        total = Salestax + Handlingfee + Shippingfee

        txtcost.Text = Cost
        txttaxdue.Text = ComputeSalesTaxDue()
        txtshipnhand.Text = ComputeTransportHandling()
        txttotal.Text = total.ToString("N1")
    End Sub


    Private Function ComputeTransportHandling()
        If (rduspostal.Checked = True) Then
            Shippingfee = 0.15
        End If

        If (rduspostalA.Checked = True) Then
            Shippingfee = 0.5
        End If

        If (txtstate.Text = "MO" Or txtstate.Text = "IL" Or txtstate.Text = "KY") Then
            Handlingfee = 0
        Else
            Handlingfee = 5
        End If
        Return Handlingfee + Shippingfee
    End Function
    Private Function ComputeSalesTaxDue()

        If (txtstate.Text = "IL") Then
            Salestax = 0.07
        End If

        If (txtstate.Text = "NY" Or txtstate.Text = "CA") Then
            Salestax = 0.08
        End If
        If (chkboxWholesale.Checked = True) Then
            Salestax = 0
        End If
        Return Salestax
    End Function

    Private Function ValidData()
        If (txtName.Text = "Null") Then
            MessageBox.Show("The Customer name can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtName.Focus()

        ElseIf (txtAddress.Text = "Null")
            MessageBox.Show("The Customer Address can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAddress.Focus()
        End If
        If (txtCity.Text = "Null") Then
            MessageBox.Show("The Customer city can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCity.Focus()

        ElseIf (txtstate.Text = "Null")
            MessageBox.Show("The Customer name can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtstate.Focus()
        End If
        If (txtZipCode.Text = "Null" Or txtZipCode.Text.Trim.Length < 5) Then
            MessageBox.Show("The ZipCode can't be Empty Or Less Than ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtZipCode.Focus()

        ElseIf (txtDescription.Text = "Null")
            MessageBox.Show("The Description can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDescription.Focus()

        End If
        If (txtPPP.Text <= 0) Then
            Try
                MessageBox.Show("The Price Per Part must be greater than zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPPP.Focus()
                txtPPP.SelectAll()
            Catch ex As Exception
            End Try
        ElseIf (IsNumeric(txtPPP.Text) = False) Then
            ValidData = False
            MessageBox.Show("The Price Per Part must be numeric", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPPP.Focus()
            txtPPP.SelectAll()


        ElseIf (Double.Parse(txtWPP.Text) <= 0) Then
            Try
                ValidData = False
                MessageBox.Show("The Weight Per Part must be greater zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtWPP.Focus()
                txtWPP.SelectAll()
            Catch ex As Exception
            End Try

        ElseIf (IsNumeric(txtWPP.Text) = False) Then
            ValidData = False
            MessageBox.Show("The Weight Per Part must be numeric", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtWPP.Focus()
            txtWPP.SelectAll()


        ElseIf (Double.Parse(txtquantity.Text) <= 0)

            Try
                ValidData = False
                MessageBox.Show("The quantity can't be Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtquantity.Focus()
                txtquantity.SelectAll()
            Catch ex As Exception

            End Try

        ElseIf (IsNumeric(txtquantity.Text) = False) Then

            MessageBox.Show("The Weight Per Part must be numeric", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtquantity.Focus()
            txtquantity.SelectAll()
        Else
            ValidData = True

        End If

    End Function

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles txtcost.TextChanged

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub
    Private Sub NewOrderbtn_Click(sender As Object, e As EventArgs) Handles NewOrderbtn.Click
        txttotal.Clear()
        txtstate.Clear()
        txtPPP.Clear()
        txtWPP.Clear()
        txtName.Clear()
        txtDescription.Clear()
        txtstate.Clear()
        txtcost.Clear()
        txtshipnhand.Clear()
        txtCity.Clear()
        txtZipCode.Clear()
        txtquantity.Clear()
        txttotal.Clear()
        chkboxWholesale.Checked = False
        rduspostal.Checked = False
        rduspostalA.Checked = False
        txtName.Focus()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Exitbtn.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtName.TextChanged

    End Sub

    Private Sub chkboxWholesale_CheckedChanged(sender As Object, e As EventArgs) Handles chkboxWholesale.CheckedChanged

    End Sub

    Private Sub txtstate_TextChanged(sender As Object, e As EventArgs) Handles txtstate.TextChanged

    End Sub

    Private Sub txtPPP_TextChanged(sender As Object, e As EventArgs) Handles txtPPP.TextChanged

    End Sub

    Private Sub txtZipCode_TextChanged(sender As Object, e As EventArgs) Handles txtZipCode.TextChanged

    End Sub

    Private Sub txtAddress_TextChanged(sender As Object, e As EventArgs) Handles txtAddress.TextChanged

    End Sub

    Private Sub txttaxdue_TextChanged(sender As Object, e As EventArgs) Handles txttaxdue.TextChanged

    End Sub
End Class
