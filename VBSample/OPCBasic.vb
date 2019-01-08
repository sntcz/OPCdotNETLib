'=======================================================================
'File:      OPCBasic.vb
'
'Summary:   sample VisualBasic.NET OPC client

'-----------------------------------------------------------------------
'  This file is part of the Viscom OPC Code Samples.
'
'  Copyright(c) 2001 Viscom (www.viscomvisual.com) All rights reserved.

'THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'PARTICULAR PURPOSE.
'========================================================================

Imports System
Imports System.Runtime.InteropServices
Imports System.Threading

Imports OPC.Common
Imports OPC.Data.Interface
Imports OPC.Data


Module TestOPCBasic

    ' ***************************************************   EDIT THIS :
    Const ServerProgID = "Softing.OPC_SMPL_SRV_ENGINE.1"    ' ProgID of OPC server
    Const ItemA = "Hauptbahnhof"                            ' fully qualified ID of a VT_I4 item
    Const ItemB = "Marienplatz"                             ' fully qualified ID of a VT_R8 item



    Public Class Tester

        Private TheSrv As OpcServer
        Private TheGrp As OpcGroup
        Private ItemDefs(1) As OPCItemDef
        Private HandlesSrv(1) As Integer


        ' -----------------------------------------------------------
        Public Sub Work()
            ' Try                       ' deactivated for debugging
            TheSrv = New OpcServer()
            TheSrv.Connect(ServerProgID)
            Thread.Sleep(500)				' we are faster then some servers!

            ' add our only working group
            TheGrp = TheSrv.AddGroup("OPCBasic-Group", False, 900)

            ' add two items and save server handles
            ItemDefs(0) = New OPCItemDef(ItemA, True, 1234, VarEnum.VT_EMPTY)
            ItemDefs(1) = New OPCItemDef(ItemB, True, 5678, VarEnum.VT_EMPTY)
            Dim rItm() As OPCItemResult
            TheGrp.AddItems(ItemDefs, rItm)
            If rItm Is Nothing Then Exit Sub
            HandlesSrv(0) = rItm(0).HandleServer
            HandlesSrv(1) = rItm(1).HandleServer

            ' asynch read our two items
            TheGrp.SetEnable(True)
            TheGrp.Active = True
            AddHandler TheGrp.DataChanged, AddressOf theGrp_DataChange
            AddHandler TheGrp.ReadCompleted, AddressOf theGrp_ReadComplete
            Dim CancelID As Integer
            Dim aE(1) As Integer
            TheGrp.Read(HandlesSrv, 55667788, CancelID, aE)

            ' some delay for asynch read-complete callback (simplification)
            Thread.Sleep(500)

            ' asynch write
            Dim ItemValues(1) As Object
            ItemValues(0) = CInt(1111111)
            ItemValues(1) = CDbl(2222.2222)
            AddHandler TheGrp.WriteCompleted, AddressOf theGrp_WriteComplete
            TheGrp.Write(HandlesSrv, ItemValues, 99887766, CancelID, aE)

            ' some delay for asynch write-complete callback (simplification)
            Thread.Sleep(500)

            ' disconnect and close
            Console.WriteLine("*********************************** hit <return> to close...")
            Console.ReadLine()
            RemoveHandler TheGrp.DataChanged, AddressOf theGrp_DataChange
            RemoveHandler TheGrp.ReadCompleted, AddressOf theGrp_ReadComplete
            RemoveHandler TheGrp.WriteCompleted, AddressOf theGrp_WriteComplete
            TheGrp.RemoveItems(HandlesSrv, aE)
            TheGrp.Remove(False)
            TheSrv.Disconnect()
            TheGrp = Nothing
            TheSrv = Nothing

            ' Catch e As Exception
            '   Console.WriteLine("VBSample: Exception")
            ' End Try
        End Sub



        ' ------------------------------ events -----------------------------

        Sub theGrp_DataChange(ByVal source As Object, ByVal e As DataChangeEventArgs)
            Console.WriteLine("DataChange event: " + e.transactionID.ToString())

            Dim s As OPCItemState
            For Each s In e.sts
                If s.Error Then
                    Console.WriteLine("  Handle:" + s.HandleClient.ToString() + " !ERROR:0x" + s.Error.ToString("X"))
                Else
                    Console.WriteLine("  Handle:" + s.HandleClient.ToString() + " Value:" + s.DataValue.ToString())
                End If
            Next
        End Sub

        Sub theGrp_ReadComplete(ByVal source As Object, ByVal e As ReadCompleteEventArgs)
            Console.WriteLine("ReadComplete event: " + e.transactionID.ToString())

            Dim s As OPCItemState
            For Each s In e.sts
                If s.Error Then
                    Console.WriteLine("  Handle:" + s.HandleClient.ToString() + " !ERROR:0x" + s.Error.ToString("X"))
                Else
                    Console.WriteLine("  Handle:" + s.HandleClient.ToString() + " Value:" + s.DataValue.ToString())
                End If
            Next
        End Sub

        Sub theGrp_WriteComplete(ByVal source As Object, ByVal e As WriteCompleteEventArgs)
            Console.WriteLine("WriteComplete event: " + e.transactionID.ToString())

            Dim r As OPCWriteResult
            For Each r In e.res
                If r.Error Then
                    Console.WriteLine("  Handle:" + r.HandleClient.ToString() + " !ERROR:0x" + r.Error.ToString("X"))
                Else
                    Console.WriteLine("  Handle:" + r.HandleClient.ToString() + " Ok.")
                End If
            Next
        End Sub

    End Class



    Sub Main()
        Console.WriteLine("OPC test VisualBasic.NET")
        Dim oT As Tester = New Tester()
        oT.Work()
    End Sub

End Module
