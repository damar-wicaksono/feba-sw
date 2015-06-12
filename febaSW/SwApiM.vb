Option Explicit On

Imports SldWorks
Imports SwConst

''' <summary>
''' VB.NET Module with functionalities to extract info from SolidWorks
''' solid model relevant to build TRACE input deck
''' </summary>
''' <remarks></remarks>
Module SwApiM

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    ''' <summary>
    ''' Function to create a discretizing box for the solid model
    ''' </summary>
    ''' <param name="swApp">The currently running Solidworks application</param>
    ''' <param name="swModel">The currently active solid models (PartDoc)</param>
    ''' <param name="dblPos">Array of Center-Face Position of the box in [m]</param>
    ''' <param name="dblDPos">Array of box dimensions in [m]</param>
    ''' <param name="blView">Flag to view the box in the model</param>
    ''' <returns>A temporary body in the model (Body2)</returns>
    ''' <remarks>Converted to VB.NET from VBA</remarks>
    Public Function createBox(ByRef swApp As SldWorks.SldWorks, _
                              ByRef swModel As SldWorks.ModelDoc2, _
                              ByVal dblPos() As Double, _
                              ByVal dblDPos() As Double, _
                              ByVal blView As Boolean) As SldWorks.Body2

        ' %-- Declare Variables
        Dim swModeler As SldWorks.Modeler
        Dim dblBoxData(8) As Double

        ' %-- Instantiate
        swModeler = swApp.GetModeler()

        ' %-- Define the box
        dblBoxData(0) = dblPos(0)   ' X-boxFaceCenter
        dblBoxData(1) = dblPos(1)   ' Y-boxFaceCenter
        dblBoxData(2) = dblPos(2)   ' Z-boxFaceCenter
        dblBoxData(3) = 0           ' X-boxAxis
        dblBoxData(4) = 1           ' Y-boxAxis
        dblBoxData(5) = 0           ' Z-boxAxis
        dblBoxData(6) = dblDPos(0)  ' boxWidth
        dblBoxData(7) = dblDPos(1)  ' boxLength
        dblBoxData(8) = dblDPos(2)  ' boxHeight

        ' %-- Create the box as temporary body
        createBox = swModeler.CreateBodyFromBox(dblBoxData)

        ' %-- View the box in SolidWorks if asked
        If blView Then
            createBox.Display3(swModel, 255, 0)
            swModel.GraphicsRedraw2()
        End If

        Return createBox
    End Function

    Public Function boxCut(ByRef swApp As SldWorks.SldWorks, _
                           ByRef swModel As SldWorks.ModelDoc2, _
                           ByVal dblPos() As Double, _
                           ByVal dblDr() As Double) As SldWorks.Body2

        Return boxCut
    End Function

    Public Sub getTraceVesselGeom()

    End Sub

    Public Function addHeatedAttr(ByVal name As String)

        ' Attribute declaration
        Dim swAttDef As SldWorks.AttributeDef
        Dim swAtt As SldWorks.Attribute
        Dim swParamName As SldWorks.Parameter
        Dim swParamValue As SldWorks.Parameter
        Dim AttName1 As String
        Dim swFeat As SldWorks.Feature
        Dim swBody As SldWorks.Body2
        Dim swEnt As SldWorks.Entity
        ' Attribute
        Const AttDefName As String = "TRACE"
        Const AttName As String = "Heated"
        Const AttValue As Integer = 1

        swApp = GetObject("", "SldWorks.Application")
        swModel = swApp.ActiveDoc()

        ' Selection by name and assign them
        swModel.Extension.SelectByID2("rods5By5<1>@AssemRods", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Dim swSelMgr As SldWorks.SelectionMgr
        swSelMgr = swModel.SelectionManager()
        swFeat = swSelMgr.GetSelectedObject6(1, 0)
        Dim objBodies As Object
        objBodies = swFeat.GetBody()
        swEnt = swFeat

        swAttDef = swApp.DefineAttribute(AttDefName)
        swAttDef.AddParameter(AttName, swParamType_e.swParamTypeString, 0.0#, 0)
        swAttDef.AddParameter(AttValue, swParamType_e.swParamTypeInteger, 0.0#, 0)
        swAttDef.Register()

        Dim i As Integer
        While swAtt Is Nothing
            i = i + 1
            AttName1 = "Damar" & CStr(i)
            swAtt = swAttDef.CreateInstance5(swModel, swEnt, AttName1, 0, swInConfigurationOpts_e.swThisConfiguration)
        End While

        swParamName = swAtt.GetParameter(AttName)
        swParamValue = swAtt.GetParameter(AttValue)

        swParamName.SetStringValue2("Damar", swInConfigurationOpts_e.swAllConfiguration, "")
        swParamValue.SetStringValue2(1, swInConfigurationOpts_e.swAllConfiguration, "")
        If Not swAtt Is Nothing Then

            Debug.Print("  " & AttDefName & "(" & i - 1 & ") = " & AttName)
            If swParamName.GetStringValue = "Damar" Then
                Debug.Print("Correct!")
            End If
        Else

            Debug.Print("  Attribute not created.")
        End If

    End Function
End Module
