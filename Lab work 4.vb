Public Class Form1

	'Lab work 4
    Private valid As Boolean = False
    Private PricePerPart As Double
    Private Quantity As Double
    Private Cost As Double
    Private SalesTax As Double = 0
    Private TaxHandling As Double = 0
    Private Handlingcharge As Double = 0

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtName.Focus()
        chkboxWholesale.Checked = False
        txtState.CharacterCasing = CharacterCasing.Upper
        radioUspostal.Checked = True
    End Sub

    Private Sub btnCompute_Click(sender As Object, e As EventArgs) Handles btnCompute.Click
        Dim total As Double
        ValidData()

        Cost = (txtPriceperpart.Text * txtQuantity.Text).ToString("n1")
        ComputeSalesTaxDue()
        ComputeTransportaionHandling()


        total = Cost + ComputeSalesTaxDue() + ComputeTransportaionHandling()

        txtCost.Text = Cost
        txtTotal.Text = total.ToString("n1")
        txtTaxDue.Text = ComputeSalesTaxDue()
        txtShipping.Text = ComputeTransportaionHandling()
    End Sub

    Private Function ComputeTransportaionHandling()
        If (radioUspostal.Checked = True) Then
            TaxHandling = 0.15
        End If

        If (radioUspostalsir.Checked = True) Then
            TaxHandling = 0.5
        End If

        If (txtState.Text = "MO" Or txtState.Text = "IL" Or txtState.Text = "KY") Then
            Handlingcharge = 0
        Else
            Handlingcharge = 5.0
        End If

        Return Handlingcharge + TaxHandling
    End Function

    Private Function ComputeSalesTaxDue()
        If (txtState.Text = "IL") Then
            SalesTax = 0.07
        End If

        If (txtState.Text = "NY" Or txtState.Text = "CA") Then
            SalesTax = 0.08
        End If

        If (chkboxWholesale.Checked = True) Then
            SalesTax = 0
        End If

        Return SalesTax
    End Function

    Private Function ValidData()

        If (txtName.Text = String.Empty) Then
            MessageBox.Show("The Customer name cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtName.Focus()

        ElseIf (txtAddress.Text = String.Empty) Then
            MessageBox.Show("The Address cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAddress.Focus()

        ElseIf (txtCity.Text = String.Empty) Then
            MessageBox.Show("The City cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCity.Focus()

        ElseIf (txtState.Text = String.Empty Or txtState.Text.Trim.Length <> 2) Then
            MessageBox.Show("The State cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtState.Focus()

        ElseIf (txtZip.Text = String.Empty Or txtZip.Text.Trim.Length < 5) Then
            MessageBox.Show("The Zip Code cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtZip.Focus()

        ElseIf (txtDescription.Text = String.Empty) Then
            MessageBox.Show("The Description cannot be empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDescription.Focus()

        ElseIf (txtPriceperpart.Text <= 0) Then
            Try

                MessageBox.Show("The Price per part must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPriceperpart.Focus()
                txtPriceperpart.SelectAll()
            Catch ex As Exception

            End Try

        ElseIf (IsNumeric(txtPriceperpart.Text) = False) Then
            valid = False
            MessageBox.Show("The Price per part must be a number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPriceperpart.Focus()
            txtPriceperpart.SelectAll()

        ElseIf (Double.Parse(txtWeightperpart.Text) <= 0) Then
            Try

                valid = False
                MessageBox.Show("The Weight per part must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtWeightperpart.Focus()
                txtWeightperpart.SelectAll()
            Catch ex As Exception

            End Try


        ElseIf (IsNumeric(txtWeightperpart.Text) = False) Then
            valid = False
            MessageBox.Show("The Weight per part must be a number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtWeightperpart.Focus()
            txtWeightperpart.SelectAll()

        ElseIf (Double.Parse(txtQuantity.Text) <= 0) Then
            Try

                valid = False
                MessageBox.Show("The Quantity must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtQuantity.Focus()
                txtQuantity.SelectAll()
            Catch ex As Exception

            End Try


        ElseIf (IsNumeric(txtQuantity.Text) = False) Then

            MessageBox.Show("The Quantity must be a number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtQuantity.Focus()
            txtQuantity.SelectAll()


        Else
            valid = True
        End If


    End Function
End Class
