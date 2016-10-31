Imports System.IO
Imports iTextSharp.text.pdf

Namespace Services.Pdf

    Public Class PdfFiller



        Private _Pdf2Fill As String = ""

        ''' <summary>
        ''' Ctor
        ''' </summary>
        ''' <param name="PDF2Fill">Chemin vers le fichier pdf � remplir</param>
        Public Sub New(ByVal pdf2Fill As String)
            _Pdf2Fill = pdf2Fill
        End Sub

        Private _Fields As New Dictionary(Of String, String)

        ''' <summary>
        ''' Dictionnaire contenant les champs � remplir avec leurs valeurs respectives
        ''' </summary>
        Public Property Fields() As Dictionary(Of String, String)
            Get
                Return _Fields
            End Get
            Set(ByVal value As Dictionary(Of String, String))
                _Fields = value
            End Set
        End Property

        ''' <summary>
        ''' Ajoute un champ � remplir
        ''' </summary>
        ''' <param name="Field">Libell� du champ � remplir</param>
        ''' <param name="Value">Valeur du champ � remplir</param>
        Public Sub AddField(ByVal field As String, ByVal value As String)
            Me._Fields(field) = value
        End Sub


        ''' <summary>
        ''' G�n�re le pdf r�sultant
        ''' </summary>
        ''' <param name="output">Stream dans lequel le pdf doit etre g�n�r�(fichier, sortie http ...)</param>
        Public Sub FillPdf(ByRef output As Stream)
            Dim ps As iTextSharp.text.pdf.PdfStamper
            Try
                '// read existing PDF document
                Dim r As New PdfReader(New RandomAccessFileOrArray(_Pdf2Fill), Nothing)
                '// optimize memory usage
                ps = New PdfStamper(r, output)
                '// retrieve properties of PDF form w/AcroFields object
                Dim af As AcroFields = ps.AcroFields
                ' // fill in PDF fields by parameter:
                '// 1. field name
                '// 2. text to insert
                For Each key As String In _Fields.Keys
                    af.SetField(key, _Fields(key))

                Next
                '// make resultant PDF read-only for end-user
                ps.FormFlattening = True
                '// forget to close() PdfStamper, you end up with
                ' // a corrupted file!
                ps.Close()
                'output.Flush()
                'output.Close()

            Catch ex As Exception
                Throw
            End Try

        End Sub
    End Class


End Namespace

