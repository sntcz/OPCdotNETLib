/*=====================================================================
  File:      OPCCSharp.cs

  Summary:   OPC sample client for C#

-----------------------------------------------------------------------
  This file is part of the Viscom OPC Code Samples.

  Copyright(c) 2001 Viscom (www.viscomvisual.com) All rights reserved.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
======================================================================*/

using System;
using System.Threading;
using System.Runtime.InteropServices;

using OPC.Common;
using OPC.Data.Interface;
using OPC.Data;

namespace CSSample
{
class Tester
	{
	// ***********************************************************	EDIT THIS :
	const string serverProgID	= "Softing.OPC_SMPL_SRV_ENGINE.1";		// ProgID of OPC server

	const string itemA			= "Hauptbahnhof";						// fully qualified ID of a VT_I4 item
	const string itemB			= "Marienplatz";						// fully qualified ID of a VT_R8 item

	private	OpcServer		theSrv;
	private	OpcGroup		theGrp;
	private	OPCItemDef[]	itemDefs = new OPCItemDef[2];
	private	int[]			handlesSrv = new int[2] { 0, 0 };


	public void Work()
		{
		/*	try						// disabled for debugging
			{	*/

			theSrv = new OpcServer();
			theSrv.Connect( serverProgID );
			Thread.Sleep(500);				// we are faster then some servers!

			// add our only working group
			theGrp = theSrv.AddGroup( "OPCCSharp-Group", false, 900 );

			// add two items and save server handles
			itemDefs[0] = new OPCItemDef( itemA, true, 1234, VarEnum.VT_EMPTY );
			itemDefs[1] = new OPCItemDef( itemB, true, 5678, VarEnum.VT_EMPTY );
			OPCItemResult[]	rItm;
			theGrp.AddItems( itemDefs, out rItm );
			if( rItm == null )
				return;
			if( HRESULTS.Failed( rItm[0].Error ) || HRESULTS.Failed( rItm[1].Error ) )
				{ Console.WriteLine( "OPC Tester: AddItems - some failed" ); theGrp.Remove( true ); theSrv.Disconnect(); return;};

			handlesSrv[0] = rItm[0].HandleServer;
			handlesSrv[1] = rItm[1].HandleServer;

			// asynch read our two items
			theGrp.SetEnable( true );
			theGrp.Active = true;
			theGrp.DataChanged += new DataChangeEventHandler( this.theGrp_DataChange );
			theGrp.ReadCompleted += new ReadCompleteEventHandler( this.theGrp_ReadComplete );
			int CancelID;
			int[] aE;
			theGrp.Read( handlesSrv, 55667788, out CancelID, out aE );

			// some delay for asynch read-complete callback (simplification)
			Thread.Sleep( 500 );


			// asynch write
			object[]	itemValues = new object[2];
			itemValues[0] = (int) 1111111;
			itemValues[1] = (double) 2222.2222;
			theGrp.WriteCompleted += new WriteCompleteEventHandler( this.theGrp_WriteComplete );
			theGrp.Write( handlesSrv, itemValues, 99887766, out CancelID, out aE );

			// some delay for asynch write-complete callback (simplification)
			Thread.Sleep( 500 );


			// disconnect and close
			Console.WriteLine( "************************************** hit <return> to close..." );
			Console.ReadLine();
			theGrp.DataChanged -= new DataChangeEventHandler( this.theGrp_DataChange );
			theGrp.ReadCompleted -= new ReadCompleteEventHandler( this.theGrp_ReadComplete );
			theGrp.WriteCompleted -= new WriteCompleteEventHandler( this.theGrp_WriteComplete );
			theGrp.RemoveItems( handlesSrv, out aE );
			theGrp.Remove( false );
			theSrv.Disconnect();
			theGrp = null;
			theSrv = null;


		/*	}
		catch( Exception e )
			{
			Console.WriteLine( "EXCEPTION : OPC Tester " + e.ToString() );
			return;
			}	*/
		}





	// ------------------------------ events -----------------------------

	public void theGrp_DataChange( object sender, DataChangeEventArgs e )
		{
		Console.WriteLine("DataChange event: gh={0} id={1} me={2} mq={3}", e.groupHandleClient, e.transactionID, e.masterError, e.masterQuality );
		foreach( OPCItemState s in e.sts )
			{
			if( HRESULTS.Succeeded( s.Error ) )
				Console.WriteLine(" ih={0} v={1} q={2} t={3}", s.HandleClient, s.DataValue, s.Quality, s.TimeStamp );
			else
				Console.WriteLine(" ih={0}    ERROR=0x{1:x} !", s.HandleClient, s.Error );
			}
		}

	public void theGrp_ReadComplete( object sender, ReadCompleteEventArgs e )
		{
		Console.WriteLine("ReadComplete event: gh={0} id={1} me={2} mq={3}", e.groupHandleClient, e.transactionID, e.masterError, e.masterQuality );
		foreach( OPCItemState s in e.sts )
			{
			if( HRESULTS.Succeeded( s.Error ) )
				Console.WriteLine(" ih={0} v={1} q={2} t={3}", s.HandleClient, s.DataValue, s.Quality, s.TimeStamp );
			else
				Console.WriteLine(" ih={0}    ERROR=0x{1:x} !", s.HandleClient, s.Error );
			}
		}

	public void theGrp_WriteComplete( object sender, WriteCompleteEventArgs e )
		{
		Console.WriteLine("WriteComplete event: gh={0} id={1} me={2}", e.groupHandleClient, e.transactionID, e.masterError );
		foreach( OPCWriteResult r in e.res )
			{
			if( HRESULTS.Succeeded( r.Error ) )
				Console.WriteLine(" ih={0} e={1}", r.HandleClient, r.Error );
			else
				Console.WriteLine(" ih={0}    ERROR=0x{1:x} !", r.HandleClient, r.Error );
			}
		}



	static void Main( string[] args )
		{
		Tester tst = new Tester();
		tst.Work();
		}
	}
}
