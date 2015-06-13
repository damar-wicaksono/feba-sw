Option Explicit On
''' <summary>
''' Module with functions/procedure for numerical calculation
''' </summary>
''' <remarks></remarks>
Module NumCrunchM

    ''' <summary>
    ''' Function to calculate the dot product of a pair 3D vector
    ''' </summary>
    ''' <param name="dblV1">Array with 3 elements</param>
    ''' <param name="dblV2">Another array with 3 elements</param>
    ''' <returns>the dot product of the two vectors</returns>
    ''' <remarks>Converted from VBA to VB.NET</remarks>
    ''' <author>
    ''' CI41, LRS/PSI, 2014
    ''' WD41, LRS/EPFL/PSI, 2015
    ''' </author>
    Public Function dotProduct(ByVal dblV1() As Double, _
                               ByVal dblV2() As Double) As Double

        dotProduct = 0.0
        For i = LBound(dblV1) To UBound(dblV1)
            dotProduct = dotProduct + dblV1(i) * dblV2(i)
        Next i

        Return dotProduct
    End Function

    ''' <summary>
    ''' Generate Z-axis cell edge for a given total length and constrained
    ''' by cell-edge and cell-centered constraints
    ''' </summary>
    ''' <param name="dblTotalLength"></param>
    ''' <param name="dblCellEdgeConstraints"></param>
    ''' <param name="dblCellCenteredConstraints"></param>
    ''' <returns>an array with edges respecting the constraints</returns>
    ''' <remarks>Adapted from the VBA Excel Macro developed by EA41</remarks>
    ''' <author>WD41, LRS/EPFL/PSI, 2015</author>
    Public Function makeEdgeZ(ByVal dblTotalLength As Double, _
                              ByVal dblCellEdgeConstraints() As Double, _
                              ByVal dblCellCenteredConstraints() As Double) _
                          As Double()

        ' %-- Variable Declarations
        Dim dblZ() As Double            ' Variable for axial position
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim intShift As Integer
        Dim dblDelt As Integer          ' Diff between Node and Edge constraints
        Dim dblPosCheck As Double
        Dim dblDeltMax As Integer
        Const dblTol As Double = 0.001  ' Tolerance limit as condition to insert new nodes

        ReDim dblZ(UBound(dblCellEdgeConstraints) + 2)

        ' %-- Assign Cell-Edge Constraints
        '     The cell-edge constraints is direct assignment ot the axial position
        ' The first edge is the bottom of the components
        dblZ(0) = 0.0#
        ' The last edge is the height of the component
        dblZ(UBound(dblZ)) = dblTotalLength
        ' Directly assign the edge constraint to the sought axial position
        For i = LBound(dblCellEdgeConstraints) To UBound(dblCellEdgeConstraints)
            dblZ(i + 1) = dblCellEdgeConstraints(i)
        Next i

        ' %-- Assign cell-centered constraints, 
        '     process additional edge if required to accomodate them
        For i = LBound(dblCellCenteredConstraints) To UBound(dblCellCenteredConstraints)

            ' Look for the edge directly below the Cell-centered constraint
            j = 1
            While dblCellCenteredConstraints(i) > dblZ(j)
                j = j + 1
            End While
            j = j - 1

            ' Compute the difference between the given node constraint
            ' and the last edge constraint
            dblDelt = dblCellCenteredConstraints(i) - dblZ(j)

            ' Check if this is the last cell-centered constraint
            If Not i = UBound(dblCellCenteredConstraints) Then
                ' Check distance is just the next cell-centered constraint
                dblPosCheck = dblCellCenteredConstraints(i + 1)
            Else
                ' Add pseudo axial position after the last cell-centered constraint
                dblPosCheck = dblCellCenteredConstraints(i) + 2 * dblDelt
            End If

            If dblCellCenteredConstraints(i) + dblDelt > dblPosCheck Or _
               dblCellCenteredConstraints(i) + dblDelt > dblZ(j + 1) Then
                ' The needed edge distance is larger than the distance of 
                ' either next cell-centered or cell-edge constraints

                ' Insert additional edge to accomodate this
                If dblCellCenteredConstraints(i) + dblDelt > dblPosCheck Then
                    ' Surpassed another node-constraint
                    dblDeltMax = (dblCellCenteredConstraints(i + 1) - _
                                 dblCellCenteredConstraints(i)) / 2
                Else
                    ' Surpassed another edge-constraint
                    dblDeltMax = dblZ(j + 1) - dblCellCenteredConstraints(i)
                End If

                intShift = 1
                If dblCellCenteredConstraints(i) - dblDeltMax - dblZ(j) > dblTol Then
                    ' The new edge if added below yields cell size larger than 
                    ' tolerance, proceed with the insertion
                    ReDim Preserve dblZ(UBound(dblZ) + 1)
                    ' Shift the all the current edge positions up to  
                    ' the new edge insertion below the constraint
                    For k = UBound(dblZ) To j + 2 Step -1
                        dblZ(k) = dblZ(k - 1)
                    Next k
                    ' Insert the new edge below the constraint
                    dblZ(j + 1) = dblCellCenteredConstraints(i) - dblDeltMax
                    intShift = 2
                End If

                If dblZ(j + intShift) - (dblCellCenteredConstraints(i) + dblDeltMax) > dblTol Then
                    ' Insert another edge above this cell-centered constraint
                    ' if appropriate
                    ReDim Preserve dblZ(UBound(dblZ) + 1)
                    ' Shift the all the current edge positions up to 
                    ' accomodate the new edge insertion above the constraint
                    For k = UBound(dblZ) To (j + intShift + 1) Step -1
                        dblZ(k) = dblZ(k - 1)
                    Next k
                    ' Insert the new edge above the constraint
                    dblZ(j + intShift) = dblCellCenteredConstraints(i) + dblDeltMax
                End If

            Else
                ' the needed edge distance is smaller than the distance of 
                ' either next cell-centered or cell-edge constraints
                If dblZ(j + 1) - (dblCellCenteredConstraints(i) + dblDelt) > dblTol Then
                    ' Insert new edge on the top of the current 
                    ' cell-centered constraint
                    ReDim Preserve dblZ(UBound(dblZ) + 1)
                    ' Shift the all the current edge positions up to 
                    ' accomodate the new edge insertion above the constraint
                    For k = UBound(dblZ) To (j + 2) Step -1
                        dblZ(k) = dblZ(k - 1)
                    Next k
                    ' Insert the new edge above the constraint
                    dblZ(j + 1) = dblCellCenteredConstraints(i) + dblDelt
                End If
            End If
        Next i

        Return dblZ
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dblEdgeZ"></param>
    ''' <param name="dblCellCenteredConstraints"></param>
    ''' <param name="dblSizeMax"></param>
    ''' <returns></returns>
    ''' <author>WD41, LRS/EPFL/PSI, 2015</author>
    Public Function refineDzbySize(ByVal dblEdgeZ() As Double, _
                                   ByVal dblCellCenteredConstraints() As Double, _
                                   ByVal dblSizeMax As Double) As Double()

        ' %-- Variable declaration
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim dblDz() As Double
        Dim intShift1 As Integer = 0    ' Index shift after insertion of new edge
        Dim intShift2 As Integer = 0    ' Index shift after node constraint was treated
        Dim intNewSize As Integer = 0   ' Integer to check if refinement is required
        Dim intOldSize As Integer = 0   ' Integer to check if refinement is required
        dblDz = discretizeZ(dblEdgeZ)
        intOldSize = 0
        intNewSize = UBound(dblDz)

        ' %-- Keep the refinement until all nodes are below dblSizeMax
        While intOldSize <> intNewSize
            intOldSize = UBound(dblDz)
            intShift1 = 0
            intShift2 = 0
            For i = LBound(dblDz) To UBound(dblDz)
                If dblDz(i) > dblSizeMax Then
                    ' Do some refinement
                    Debug.Print("Node Size " & i & "= " & dblDz(i) & " Refine Node")
                    For j = intShift2 To UBound(dblCellCenteredConstraints)

                        If dblCellCenteredConstraints(j) > dblEdgeZ(i + intShift1) And dblCellCenteredConstraints(j) < dblEdgeZ(i + intShift1 + 1) Then
                            ' There is node constraint in this node, divide by 3
                            'Debug.Print("Node constraint exists")
                            ' Increase edge array and shift up
                            ReDim Preserve dblEdgeZ(UBound(dblEdgeZ) + 2)
                            'For k = UBound(dblEdgeZ) To (i + intShift1 + 3) Step -1
                            '    dblEdgeZ(k) = dblEdgeZ(k - 2)
                            '    Debug.Print(dblEdgeZ(k))
                            'Next k
                            ' Insert two new edges
                            dblEdgeZ(i + intShift1 + 1) = dblEdgeZ(i + intShift1) + dblDz(i) / 3
                            dblEdgeZ(i + intShift1 + 2) = dblEdgeZ(i + intShift1) + 2 * dblDz(i) / 3
                            intShift1 = intShift1 + 2
                            'Debug.Print("Constraint Exist, new edge at " & dblEdgeZ(i + intShift1 - 1))
                            'Debug.Print("Constraint Exist, new edge at " & dblEdgeZ(i + intShift1))
                            intShift2 = intShift2 + 1
                            'For k = LBound(dblEdgeZ) To UBound(dblEdgeZ)
                            '    Debug.Print(dblEdgeZ(k))
                            'Next k
                            Exit For
                        Else
                            ' No node constraint, divide by 2
                            ' Increase edge array and shift up
                            ReDim Preserve dblEdgeZ(UBound(dblEdgeZ) + 1)
                            For k = UBound(dblEdgeZ) To (i + intShift1 + 2) Step -1
                                dblEdgeZ(k) = dblEdgeZ(k - 1)
                            Next k
                            'Insert new edge
                            'dblDelt = (dblEdgeZ(i + 1) - dblEdgeZ(i)) / 2
                            dblEdgeZ(i + intShift1 + 1) = dblEdgeZ(i + intShift1) + dblDz(i) / 2
                            intShift1 = intShift1 + 1
                            'Debug.Print("No Constraint, new edge at " & dblEdgeZ(i + intShift1))
                            'For k = LBound(dblEdgeZ) To UBound(dblEdgeZ)
                            '    Debug.Print(dblEdgeZ(k))
                            'Next k
                            Exit For
                        End If
                    Next j
                End If
            Next i
            dblDz = discretizeZ(dblEdgeZ)
            intNewSize = UBound(dblDz)
        End While

        Return dblEdgeZ
    End Function

    ''' <summary>
    ''' Create an array of cell size based on edge positions
    ''' </summary>
    ''' <param name="dblEdgeZ">Array of edge positions</param>
    ''' <returns>Array of cell sizes</returns>
    ''' <author>WD41, LRS/EPFL/PSI 2015</author>
    Public Function discretizeZ(ByVal dblEdgeZ() As Double) As Double()

        ' %-- Variable declaration
        Dim dblDz() As Double           ' Variable for cell size 
        Dim i As Integer

        ' %-- Compute the cell size for a given set of edges
        ReDim dblDz(UBound(dblEdgeZ) - 1)
        For i = LBound(dblDz) To UBound(dblDz)
            dblDz(i) = dblEdgeZ(i + 1) - dblEdgeZ(i)
        Next i

        Return dblDz
    End Function

    Public Function getCellCenteredZ(ByVal dblEdgeZ() As Double) As Double()

        Dim dblCellCenteredZ() As Double
        Dim i As Integer
        Dim dblDelt As Double
        ReDim dblCellCenteredZ(UBound(dblEdgeZ) - 1)

        For i = LBound(dblCellCenteredZ) To UBound(dblCellCenteredZ)
            dblDelt = dblEdgeZ(i + 1) - dblEdgeZ(i)
            dblCellCenteredZ(i) = dblEdgeZ(i) + dblDelt / 2
        Next i

        Return dblCellCenteredZ
    End Function

End Module
