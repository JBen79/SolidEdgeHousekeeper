Option Strict On

Public Class TaskFinalizeModelAndDrawings

    Inherits Task

    Public Sub New()
        Me.Name = Me.ToString.Replace("Housekeeper.", "")
        Me.Description = "Finalize model and drawings"
        Me.HelpText = GetHelpText()
        Me.RequiresSave = True
        Me.AppliesToAssembly = True
        Me.AppliesToPart = True
        Me.AppliesToSheetmetal = True
        Me.AppliesToDraft = True
        Me.HasOptions = False
        Me.HelpURL = GenerateHelpURL("Finalize model and drawings")
        Me.Image = My.Resources.TaskFitView
        Me.Category = "Restyle"
        SetColorFromCategory(Me)
        Me.RequiresForegroundProcessing = True

        GenerateTaskControl()
    End Sub


    Public Overrides Sub Process(
        ByVal SEDoc As SolidEdgeFramework.SolidEdgeDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        Me.TaskLogger = Me.FileLogger.AddLogger(Me.Description)

        InvokeSTAThread(
            Of SolidEdgeFramework.SolidEdgeDocument,
            SolidEdgeFramework.Application)(
                AddressOf ProcessInternal,
                SEDoc,
                SEApp)
    End Sub


    Public Overrides Sub Process(ByVal FileName As String)
        Me.TaskLogger = Me.FileLogger.AddLogger(Me.Description)
    End Sub


    Private Sub ProcessInternal(
        ByVal SEDoc As SolidEdgeFramework.SolidEdgeDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        Dim UC As New UtilsCommon
        Dim DocType As String = UC.GetDocType(SEDoc)

        Select Case DocType
            Case "asm"
                HideAllModelConstructions(SEDoc, SEApp, DocType)
                OrientAssemblyIsometric(SEApp)
                UpdateAssemblyPhysicalProperties(
                    CType(SEDoc, SolidEdgeAssembly.AssemblyDocument),
                    SEApp)

            Case "par", "psm"
                HideAllModelConstructions(SEDoc, SEApp, DocType)
                OrientPartOrSheetMetalIsometric(SEApp)
                UpdatePartOrSheetMetalPhysicalProperties(SEDoc, SEApp, DocType)

            Case "dft"
                UpdateAndFitDraft(
                    CType(SEDoc, SolidEdgeDraft.DraftDocument),
                    SEApp)

            Case Else
                TaskLogger.AddMessage(
                    String.Format("Document type '{0}' not recognized", DocType))
                Return
        End Select

        If SEDoc.ReadOnly Then
            TaskLogger.AddMessage("Cannot save document marked 'Read Only'")
        Else
            Try
                SEDoc.Save()
                SEApp.DoIdle()
            Catch ex As Exception
                TaskLogger.AddMessage("Unable to save document")
            End Try
        End If
    End Sub


    Private Sub HideAllModelConstructions(
        ByVal SEDoc As SolidEdgeFramework.SolidEdgeDocument,
        ByVal SEApp As SolidEdgeFramework.Application,
        ByVal DocType As String)

        ' These commands match the existing Show/Hide constructions task.
        ' The sketch commands also hide 3D sketches.
        Select Case DocType
            Case "asm"
                HideAllAssemblyReferencePlanes(
                    CType(SEDoc, SolidEdgeAssembly.AssemblyDocument),
                    SEApp)
                ExecuteDisplayCommand(SEApp, 40080, "Unable to hide all sketches")
                ExecuteDisplayCommand(SEApp, 40081, "Unable to hide all coordinate systems")

            Case "par", "psm"
                ExecuteDisplayCommand(SEApp, 40229, "Unable to hide all reference planes")
                ExecuteDisplayCommand(SEApp, 40252, "Unable to hide all sketches")
                ExecuteDisplayCommand(SEApp, 40270, "Unable to hide all coordinate systems")
        End Select
    End Sub


    Private Sub HideAllAssemblyReferencePlanes(
        ByVal AssemblyDocument As SolidEdgeAssembly.AssemblyDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        Try
            ' Empty assemblies use the part-style command; populated assemblies use
            ' the dedicated assembly command that hides all reference planes.
            If AssemblyDocument.Occurrences.Count = 0 Then
                SEApp.StartCommand(
                    CType(40229, SolidEdgeFramework.SolidEdgeCommandConstants))
            Else
                SEApp.StartCommand(
                    CType(
                        SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsHideAllReferencePlanes,
                        SolidEdgeFramework.SolidEdgeCommandConstants))
            End If

            SEApp.DoIdle()
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to hide all reference planes")
        End Try
    End Sub


    Private Sub ExecuteDisplayCommand(
        ByVal SEApp As SolidEdgeFramework.Application,
        ByVal CommandId As Integer,
        ByVal ErrorMessage As String)

        Try
            SEApp.StartCommand(
                CType(CommandId, SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()
        Catch ex As Exception
            TaskLogger.AddMessage(ErrorMessage)
        End Try
    End Sub


    Private Sub OrientAssemblyIsometric(
        ByVal SEApp As SolidEdgeFramework.Application)

        Try
            SEApp.StartCommand(
                CType(
                    SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView,
                    SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()

            SEApp.StartCommand(
                CType(
                    SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit,
                    SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to orient assembly to the isometric view")
        End Try
    End Sub


    Private Sub OrientPartOrSheetMetalIsometric(
        ByVal SEApp As SolidEdgeFramework.Application)

        Try
            SEApp.StartCommand(
                CType(
                    SolidEdgeConstants.PartCommandConstants.PartViewISOView,
                    SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()

            SEApp.StartCommand(
                CType(
                    SolidEdgeConstants.PartCommandConstants.PartViewFit,
                    SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to orient model to the isometric view")
        End Try
    End Sub


    Private Sub UpdateAssemblyPhysicalProperties(
        ByVal AssemblyDocument As SolidEdgeAssembly.AssemblyDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        Try
            If AssemblyDocument.Occurrences.Count = 0 Then
                Return
            End If

            Dim PhysicalProperties As SolidEdgeAssembly.PhysicalProperties
            PhysicalProperties = AssemblyDocument.PhysicalProperties

            ' Use the same API call used by Housekeeper's native
            ' Update physical properties task.
            Dim FilesWithoutDensityArray As Array = {""}

            PhysicalProperties.UpdateV2(FilesWithoutDensityArray)
            SEApp.DoIdle()

            Dim MissingDensityCount As Integer = 0

            If FilesWithoutDensityArray IsNot Nothing Then
                For Each Item As Object In FilesWithoutDensityArray
                    Dim FileName As String = Convert.ToString(Item)

                    If Not String.IsNullOrWhiteSpace(FileName) Then
                        MissingDensityCount += 1
                    End If
                Next
            End If

            If MissingDensityCount > 0 Then
                TaskLogger.AddMessage(
                    String.Format(
                        "Found {0} models with no density assigned. Please verify results.",
                        MissingDensityCount))
            End If

        Catch ex As Exception
            TaskLogger.AddMessage(
                String.Format(
                    "Unable to update assembly physical properties: {0}",
                    ex.Message))
        End Try
    End Sub


    Private Sub UpdatePartOrSheetMetalPhysicalProperties(
        ByVal SEDoc As SolidEdgeFramework.SolidEdgeDocument,
        ByVal SEApp As SolidEdgeFramework.Application,
        ByVal DocType As String)

        Dim Models As SolidEdgePart.Models = Nothing
        Dim Density As Double = 0.0

        Try
            Dim MaterialTable As SolidEdgeFramework.MatTable
            MaterialTable = SEApp.GetMaterialTable()

            Dim PropertyType As SolidEdgeFramework.MatTablePropIndexConstants
            PropertyType = SolidEdgeFramework.MatTablePropIndexConstants.seDensity

            Dim PropertyValue As Object = Nothing

            If DocType = "par" Then
                Dim PartDocument As SolidEdgePart.PartDocument
                PartDocument = CType(SEDoc, SolidEdgePart.PartDocument)
                Models = PartDocument.Models

                MaterialTable.GetMaterialPropValueFromDoc(
                    PartDocument,
                    CType(PropertyType, SolidEdgeFramework.MatTablePropIndex),
                    PropertyValue)
            Else
                Dim SheetMetalDocument As SolidEdgePart.SheetMetalDocument
                SheetMetalDocument = CType(SEDoc, SolidEdgePart.SheetMetalDocument)
                Models = SheetMetalDocument.Models

                MaterialTable.GetMaterialPropValueFromDoc(
                    SheetMetalDocument,
                    CType(PropertyType, SolidEdgeFramework.MatTablePropIndex),
                    PropertyValue)
            End If

            If PropertyValue IsNot Nothing Then
                Density = CDbl(PropertyValue)
            End If

        Catch ex As Exception
            TaskLogger.AddMessage(
                String.Format(
                    "Unable to read the model density: {0}",
                    ex.Message))
            Return
        End Try

        If Models Is Nothing OrElse Models.Count = 0 Then
            TaskLogger.AddMessage(
                "No solid model found for physical property update")
            Return
        End If

        If Density <= 0 Then
            TaskLogger.AddMessage(
                String.Format("Density set to {0}", Density))
            Return
        End If

        Try
            ' Use the same Solid Edge command used by Housekeeper's native
            ' Update physical properties task. This avoids COM array type
            ' mismatches in ComputePhysicalProperties.
            SEApp.StartCommand(
                CType(
                    SolidEdgeConstants.PartCommandConstants.PartToolsPhysicalProperties,
                    SolidEdgeFramework.SolidEdgeCommandConstants))
            SEApp.DoIdle()

        Catch ex As Exception
            TaskLogger.AddMessage(
                String.Format(
                    "Unable to update model physical properties: {0}",
                    ex.Message))
        End Try
    End Sub


    Private Sub UpdateAndFitDraft(
        ByVal DraftDocument As SolidEdgeDraft.DraftDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        UpdateAllDrawingViews(DraftDocument)
        UpdateAllDraftTables(DraftDocument)

        Try
            DraftDocument.UpdatePropertyTextCacheAndDisplay()
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to refresh property text")
        End Try

        FitAllWorkingSheets(DraftDocument, SEApp)
    End Sub


    Private Sub UpdateAllDrawingViews(
        ByVal DraftDocument As SolidEdgeDraft.DraftDocument)

        Dim UC As New UtilsCommon

        For Each Sheet As SolidEdgeDraft.Sheet In UC.GetSheets(DraftDocument, "Working")
            For Each DrawingView As SolidEdgeDraft.DrawingView In
                Sheet.DrawingViews.OfType(Of SolidEdgeDraft.DrawingView)()

                Try
                    ' ForceUpdate updates the view even when Solid Edge reports it as current.
                    DrawingView.ForceUpdate()
                Catch exForce As Exception
                    Try
                        DrawingView.Update()
                    Catch exUpdate As Exception
                        TaskLogger.AddMessage(
                            String.Format(
                                "Unable to update drawing view on sheet '{0}'",
                                Sheet.Name))
                    End Try
                End Try
            Next
        Next
    End Sub


    Private Sub UpdateAllDraftTables(
        ByVal DraftDocument As SolidEdgeDraft.DraftDocument)

        Try
            For Each PartsList As SolidEdgeDraft.PartsList In DraftDocument.PartsLists
                Try
                    PartsList.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update parts list")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access parts lists")
        End Try

        Try
            For Each HoleTable As SolidEdgeDraft.HoleTable In DraftDocument.HoleTables
                Try
                    HoleTable.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update hole table")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access hole tables")
        End Try

        Try
            For Each BendTable As SolidEdgeDraft.DraftBendTable In DraftDocument.DraftBendTables
                Try
                    BendTable.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update bend table")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access bend tables")
        End Try

        Try
            For Each BlockTable As SolidEdgeDraft.BlockTable In DraftDocument.BlockTables
                Try
                    BlockTable.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update block table")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access block tables")
        End Try

        Try
            For Each ConnectorTable As SolidEdgeDraft.ConnectorTable In DraftDocument.ConnectorTables
                Try
                    ConnectorTable.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update connector table")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access connector tables")
        End Try

        Try
            For Each UserTable As SolidEdgeDraft.Table In DraftDocument.Tables
                Try
                    UserTable.Update()
                Catch ex As Exception
                    TaskLogger.AddMessage("Unable to update user table")
                End Try
            Next
        Catch ex As Exception
            TaskLogger.AddMessage("Unable to access user tables")
        End Try
    End Sub


    Private Sub FitAllWorkingSheets(
        ByVal DraftDocument As SolidEdgeDraft.DraftDocument,
        ByVal SEApp As SolidEdgeFramework.Application)

        Try
            Dim WorkingSheets As SolidEdgeDraft.SectionSheets
            WorkingSheets = DraftDocument.Sections.WorkingSection.Sheets

            If WorkingSheets.Count = 0 Then
                TaskLogger.AddMessage("No working sheets found")
                Return
            End If

            Dim SheetWindow As SolidEdgeDraft.SheetWindow
            SheetWindow = CType(SEApp.ActiveWindow, SolidEdgeDraft.SheetWindow)

            For Each Sheet As SolidEdgeDraft.Sheet In
                WorkingSheets.OfType(Of SolidEdgeDraft.Sheet)()

                SheetWindow.ActiveSheet = Sheet
                SheetWindow.FitEx(SolidEdgeDraft.SheetFitConstants.igFitSheet)
                SheetWindow.Update()
                SEApp.DoIdle()
            Next

            ' Leave the first working sheet visible when processing is complete.
            SheetWindow.ActiveSheet =
                WorkingSheets.OfType(Of SolidEdgeDraft.Sheet)().ElementAt(0)
            SheetWindow.Update()
            SEApp.DoIdle()

        Catch ex As Exception
            TaskLogger.AddMessage("Unable to fit all drawing sheets")
        End Try
    End Sub


    Public Overrides Sub CheckStartConditions(ErrorLogger As Logger)

        If Me.IsSelectedTask Then
            If Not (
                Me.IsSelectedAssembly OrElse
                Me.IsSelectedPart OrElse
                Me.IsSelectedSheetmetal OrElse
                Me.IsSelectedDraft) Then

                ErrorLogger.AddMessage("Select at least one type of file to process")
            End If
        End If
    End Sub


    Private Function GetHelpText() As String
        Dim HelpString As String

        HelpString = "For part, sheet metal, and assembly documents, hides all sketches, reference planes, and coordinate systems, sets the active window to the isometric view, and fits the model to the window. "
        HelpString += "It updates physical properties for parts and sheet metal files, and updates assembly physical properties. "
        HelpString += "For draft documents, it updates all drawing views and supported tables, fits every working sheet, leaves the first sheet active, and saves the document."

        Return HelpString
    End Function

End Class
