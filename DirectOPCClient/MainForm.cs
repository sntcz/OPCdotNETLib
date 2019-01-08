using System;
using System.Threading;
using System.Text;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;

using OPC.Common;
using OPC.Data.Interface;
using OPC.Data;

namespace DirectOPCClient
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.ComponentModel.IContainer components;

		private System.Windows.Forms.TextBox txtServer;
		private System.Windows.Forms.TextBox txtServerInfo;
		private System.Windows.Forms.TreeView treeOpcItems;
		private System.Windows.Forms.ListView listOpcView;
		private System.Windows.Forms.TextBox txtItemID;
		private System.Windows.Forms.TextBox txtItemQual;
		private System.Windows.Forms.TextBox txtItemTimeSt;
		private System.Windows.Forms.TextBox txtItemValue;
		private System.Windows.Forms.TextBox txtItemDataType;
		private System.Windows.Forms.TextBox txtItemSendValue;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ColumnHeader hdrListCol1;
		private System.Windows.Forms.ColumnHeader hdrListCol2;
		private System.Windows.Forms.Button btnItemWrite;		// workaround to show SelServer Form on applic. start
		private System.Windows.Forms.StatusBar statusBarMain;
		private System.Windows.Forms.StatusBarPanel sbpTimeStart;
		private System.Windows.Forms.StatusBarPanel sbpStatus;
		private System.Windows.Forms.StatusBarPanel sbpDummy;				// node in TreeView
		private System.Windows.Forms.ImageList opctreeicons;
		private System.Windows.Forms.ImageList opclisticons;
		private System.Windows.Forms.TextBox txtItemWriteRes;
		private System.Windows.Forms.Button btnItemMore;
		private System.Windows.Forms.Button btnAbout;


		public Process			thisprocess;			// running OS process

		public string			selectedOpcSrv;			// Name (ProgID) of selected OPC-server
		public OpcServer		theSrv = null;			// root OPCDA object
		public OpcGroup			theGrp = null;			// the only one OPC-Group in this example

		public string			itmFullID;					// fully qualified OPC namespace path
		public int				itmHandleClient;			// 0 if no current item selected
		public int				itmHandleServer;
		public OPCACCESSRIGHTS	itmAccessRights;
		public TypeCode			itmTypeCode;				// saved data type of current item
	
		public bool				first_activated = false;	// workaround to show SelServer Form on applic. start
		public bool				opc_connected = false;		// flag if connected

		public string			rootname = "Root";			// string of TreeView root (dummy)
		public string			selectednode;
		public string			selecteditem;				// item in ListView


		public MainForm()
			{
			thisprocess = Process.GetCurrentProcess();		// see DoConnect for client-name
			InitializeComponent();

			treeOpcItems.PathSeparator = "\t";				// warning: assuming OPC not using tabulator as separator
			}

	protected void theSrv_ServerShutDown( object sender, ShutdownRequestEventArgs e )
		{					// event: the OPC server shuts down
		MessageBox.Show( this, "OPC server shuts down because:" + e.shutdownReason, "ServerShutDown", MessageBoxButtons.OK, MessageBoxIcon.Warning );
		}

	public bool DoInit()
		{
		try
			{
			SelServer	frmSelSrv = new SelServer( );		// create form and let user select a name
			frmSelSrv.ShowDialog( this );
			if( frmSelSrv.selectedOpcSrv == null )
				this.Close();

			selectedOpcSrv = frmSelSrv.selectedOpcSrv;			// OPC server ProgID
			txtServer.Text = selectedOpcSrv;


			// ---------------
			theSrv = new OpcServer();
			if( ! DoConnect( selectedOpcSrv ) )
				return false;

			// add event handler for server shutdown
			theSrv.ShutdownRequested += new ShutdownRequestEventHandler( this.theSrv_ServerShutDown );

			// precreate the only OPC group in this example
			if( ! CreateGroup() )
				return false;

			// browse the namespace of the OPC-server
			if( ! DoBrowse() )
				return false;
			}
		catch( Exception e )		// exceptions MUST be handled
			{
			MessageBox.Show( this, "init error! " + e.ToString(), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}

	// connect to OPC server via ProgID
	public bool DoConnect( string progid )
		{
		try
			{
			theSrv.Connect( progid );
			Thread.Sleep( 100 );
			theSrv.SetClientName( "DirectOPC " + thisprocess.Id );	// set my client name (exe+process no)

			SERVERSTATUS sts;
			theSrv.GetStatus( out sts );

			// get infos about OPC server
			StringBuilder sb = new StringBuilder( sts.szVendorInfo, 200 );
			sb.AppendFormat( " ver:{0}.{1}.{2}", sts.wMajorVersion, sts.wMinorVersion, sts.wBuildNumber );
			txtServerInfo.Text = sb.ToString();

			// set status bar text to show server state
			sbpTimeStart.Text = DateTime.FromFileTime( sts.ftStartTime ).ToString();
			sbpStatus.Text = sts.eServerState.ToString();
			}
		catch( COMException )
			{
			MessageBox.Show( this, "connect error!", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}


	public bool CreateGroup()
		{
		try
			{
			// add our only working group
			theGrp = theSrv.AddGroup( "OPCdotNET-Group", true, 500 );

			// add event handler for data changes
			theGrp.DataChanged += new DataChangeEventHandler( this.theGrp_DataChange );
			theGrp.WriteCompleted += new WriteCompleteEventHandler( this.theGrp_WriteComplete );
			}
		catch( COMException )
			{
			MessageBox.Show( this, "create group error!", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}

	// event handler: called if any item in group has changed values
	protected void theGrp_DataChange( object sender, DataChangeEventArgs e )
		{
		Trace.WriteLine( "theGrp_DataChange  id=" + e.transactionID.ToString() + " me=0x" + e.masterError.ToString( "X" ) );

		foreach( OPCItemState s in e.sts )
			{
			if( s.HandleClient != itmHandleClient )		// only one client handle
				continue;

			Trace.WriteLine( "  item error=0x" + s.Error.ToString( "X" ) );

			if( HRESULTS.Succeeded( s.Error ) )
				{
				Trace.WriteLine( "  val=" + s.DataValue.ToString() );

				txtItemValue.Text	= s.DataValue.ToString();		// update screen
				txtItemQual.Text	= OpcGroup.QualityToString( s.Quality );
				txtItemTimeSt.Text	= DateTime.FromFileTime( s.TimeStamp ).ToString();
				}
			else
				{
				txtItemValue.Text	= "ERROR 0x" + s.Error.ToString( "X" );
				txtItemQual.Text	= "error";
				txtItemTimeSt.Text	= "error";
				}
			}
		}

	// event handler: called if asynch write finished
	protected void theGrp_WriteComplete( object sender, WriteCompleteEventArgs e )
		{
		foreach( OPCWriteResult w in e.res )
			{
			if( w.HandleClient != itmHandleClient )		// only one client handle
				continue;

			if( HRESULTS.Failed( w.Error ) )
				txtItemWriteRes.Text	= "ERROR 0x" + w.Error.ToString( "X" );
			else
				txtItemWriteRes.Text	= "ok";
			}
		}


	public bool DoBrowse()
		{
		try
			{
			OPCNAMESPACETYPE	opcorgi = theSrv.QueryOrganization();

			// fill TreeView with all
			treeOpcItems.Nodes.Clear();
			TreeNode	tnRoot = new TreeNode( rootname, 0, 1 );
			if( opcorgi == OPCNAMESPACETYPE.OPC_NS_HIERARCHIAL )
				{
				theSrv.ChangeBrowsePosition( OPCBROWSEDIRECTION.OPC_BROWSE_TO, "" );	// to root
				RecurBrowse( tnRoot, 1 );
				}
			treeOpcItems.Nodes.Add( tnRoot );

			tnRoot.ExpandAll();			// expand all nodes ([+] -> [-])
			tnRoot.EnsureVisible();		// make the root visible

			// preselect root (dummy)
			treeOpcItems.SelectedNode = tnRoot;		// force treeOpcItems_AfterSelect
			}
		catch( COMException /* eX */ )
			{
			MessageBox.Show( this, "browse error!", "DoBrowse", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}

	// recursively call the OPC namespace tree
	public bool RecurBrowse( TreeNode tnParent, int depth )
		{
		try
			{
			ArrayList lst;
			theSrv.Browse( OPCBROWSETYPE.OPC_BRANCH, out lst );
			if( lst == null )
				return true;
			if( lst.Count < 1 )
				return true;
		
			foreach( string s in lst )
				{
				TreeNode tnNext = new TreeNode( s, 0, 1 );
				theSrv.ChangeBrowsePosition( OPCBROWSEDIRECTION.OPC_BROWSE_DOWN, s );
				RecurBrowse( tnNext, depth + 1 );
				theSrv.ChangeBrowsePosition( OPCBROWSEDIRECTION.OPC_BROWSE_UP, "" );
			
				tnParent.Nodes.Add( tnNext );
				}
			}
		catch( COMException /* eX */ )
			{
			MessageBox.Show( this, "browse error!", "RecurBrowse", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}


	public bool ViewItem( string opcid )
		{
		try
			{
			RemoveItem();		// first remove previous item if any

			itmHandleClient	= 1234;
			OPCItemDef[]	aD = new OPCItemDef[1];
			aD[0] = new OPCItemDef( opcid, true, itmHandleClient, VarEnum.VT_EMPTY );
			OPCItemResult[]		arrRes;
			theGrp.AddItems( aD, out arrRes );
			if( arrRes == null )
				return false;
			if( arrRes[0].Error != HRESULTS.S_OK )
				return false;

			btnItemMore.Enabled	= true;
			itmHandleServer		= arrRes[0].HandleServer;
			itmAccessRights		= arrRes[0].AccessRights;
			itmTypeCode			= VT2TypeCode( arrRes[0].CanonicalDataType );

			txtItemID.Text			= opcid;
			txtItemDataType.Text	= DUMMY_VARIANT.VarEnumToString( arrRes[0].CanonicalDataType );

			if( (itmAccessRights & OPCACCESSRIGHTS.OPC_READABLE) != 0 )
				{
				int		cancelID;
				theGrp.Refresh2( OPCDATASOURCE.OPC_DS_DEVICE, 7788, out cancelID );
				}
			else
				txtItemValue.Text = "no read access";

			if( itmTypeCode != TypeCode.Object )				// Object=failed!
				{
				// check if write is premitted
				if( (itmAccessRights & OPCACCESSRIGHTS.OPC_WRITEABLE) != 0 )
					btnItemWrite.Enabled = true;
				}
			}
		catch( COMException )
			{
			MessageBox.Show( this, "AddItem OPC error!", "ViewItem", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}


	// remove previous OPC item if any
	public bool RemoveItem()
		{
		try
			{
			if( itmHandleClient != 0 )
				{
				itmHandleClient			= 0;
				txtItemID.Text			= "";		// clear screen texts
				txtItemValue.Text		= "";
				txtItemDataType.Text	= "";
				txtItemQual.Text		= "";
				txtItemTimeSt.Text		= "";
				txtItemSendValue.Text	= "";
				txtItemWriteRes.Text	= "";
				btnItemWrite.Enabled	= false;
				btnItemMore.Enabled		= false;

				int[]	serverhandles	= new int[1] { itmHandleServer };
				int[]	remerrors;
				theGrp.RemoveItems( serverhandles, out remerrors );
				itmHandleServer = 0;
				}
			}
		catch( COMException )
			{
			MessageBox.Show( this, "RemoveItem OPC error!", "RemoveItem", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return false;
			}
		return true;
		}





	/// <summary>
	/// Clean up any resources being used.
	/// </summary>
	protected override void Dispose( bool disposing )
		{
		if( disposing )
			{
			if (components != null) 
				{
				components.Dispose();
				}
			}
		base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.txtItemDataType = new System.Windows.Forms.TextBox();
			this.btnItemWrite = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.txtServer = new System.Windows.Forms.TextBox();
			this.txtItemID = new System.Windows.Forms.TextBox();
			this.listOpcView = new System.Windows.Forms.ListView();
			this.hdrListCol1 = new System.Windows.Forms.ColumnHeader();
			this.hdrListCol2 = new System.Windows.Forms.ColumnHeader();
			this.opclisticons = new System.Windows.Forms.ImageList(this.components);
			this.txtServerInfo = new System.Windows.Forms.TextBox();
			this.sbpStatus = new System.Windows.Forms.StatusBarPanel();
			this.btnAbout = new System.Windows.Forms.Button();
			this.txtItemValue = new System.Windows.Forms.TextBox();
			this.txtItemWriteRes = new System.Windows.Forms.TextBox();
			this.opctreeicons = new System.Windows.Forms.ImageList(this.components);
			this.txtItemTimeSt = new System.Windows.Forms.TextBox();
			this.statusBarMain = new System.Windows.Forms.StatusBar();
			this.sbpTimeStart = new System.Windows.Forms.StatusBarPanel();
			this.sbpDummy = new System.Windows.Forms.StatusBarPanel();
			this.btnItemMore = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.treeOpcItems = new System.Windows.Forms.TreeView();
			this.txtItemSendValue = new System.Windows.Forms.TextBox();
			this.txtItemQual = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.sbpStatus)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpTimeStart)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpDummy)).BeginInit();
			this.SuspendLayout();
			// 
			// txtItemDataType
			// 
			this.txtItemDataType.Location = new System.Drawing.Point(336, 296);
			this.txtItemDataType.Name = "txtItemDataType";
			this.txtItemDataType.ReadOnly = true;
			this.txtItemDataType.Size = new System.Drawing.Size(184, 20);
			this.txtItemDataType.TabIndex = 8;
			this.txtItemDataType.Text = "";
			// 
			// btnItemWrite
			// 
			this.btnItemWrite.Enabled = false;
			this.btnItemWrite.Location = new System.Drawing.Point(296, 328);
			this.btnItemWrite.Name = "btnItemWrite";
			this.btnItemWrite.Size = new System.Drawing.Size(64, 24);
			this.btnItemWrite.TabIndex = 11;
			this.btnItemWrite.Text = "write!";
			this.btnItemWrite.Click += new System.EventHandler(this.btnItemWrite_Click);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 276);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 12;
			this.label3.Text = "quality";
			this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// txtServer
			// 
			this.txtServer.Location = new System.Drawing.Point(8, 8);
			this.txtServer.Name = "txtServer";
			this.txtServer.ReadOnly = true;
			this.txtServer.Size = new System.Drawing.Size(248, 20);
			this.txtServer.TabIndex = 0;
			this.txtServer.Text = "server";
			// 
			// txtItemID
			// 
			this.txtItemID.Location = new System.Drawing.Point(8, 240);
			this.txtItemID.Name = "txtItemID";
			this.txtItemID.ReadOnly = true;
			this.txtItemID.Size = new System.Drawing.Size(512, 20);
			this.txtItemID.TabIndex = 4;
			this.txtItemID.Text = "full";
			// 
			// listOpcView
			// 
			this.listOpcView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.hdrListCol1,
																						  this.hdrListCol2});
			this.listOpcView.FullRowSelect = true;
			this.listOpcView.GridLines = true;
			this.listOpcView.HideSelection = false;
			this.listOpcView.Location = new System.Drawing.Point(264, 32);
			this.listOpcView.MultiSelect = false;
			this.listOpcView.Name = "listOpcView";
			this.listOpcView.Size = new System.Drawing.Size(256, 200);
			this.listOpcView.SmallImageList = this.opclisticons;
			this.listOpcView.TabIndex = 3;
			this.listOpcView.View = System.Windows.Forms.View.Details;
			this.listOpcView.SelectedIndexChanged += new System.EventHandler(this.listOpcView_SelectedIndexChanged);
			// 
			// hdrListCol1
			// 
			this.hdrListCol1.Text = "Items";
			this.hdrListCol1.Width = 128;
			// 
			// hdrListCol2
			// 
			this.hdrListCol2.Text = "ItemID";
			this.hdrListCol2.Width = 300;
			// 
			// opclisticons
			// 
			this.opclisticons.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
			this.opclisticons.ImageSize = new System.Drawing.Size(16, 16);
			this.opclisticons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("opclisticons.ImageStream")));
			this.opclisticons.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtServerInfo
			// 
			this.txtServerInfo.Location = new System.Drawing.Point(264, 8);
			this.txtServerInfo.Name = "txtServerInfo";
			this.txtServerInfo.ReadOnly = true;
			this.txtServerInfo.Size = new System.Drawing.Size(256, 20);
			this.txtServerInfo.TabIndex = 1;
			this.txtServerInfo.Text = "serverinfo";
			// 
			// sbpStatus
			// 
			this.sbpStatus.MinWidth = 32;
			this.sbpStatus.Text = "1";
			this.sbpStatus.ToolTipText = "status of OPC server";
			this.sbpStatus.Width = 192;
			// 
			// btnAbout
			// 
			this.btnAbout.BackColor = System.Drawing.Color.SandyBrown;
			this.btnAbout.Location = new System.Drawing.Point(456, 360);
			this.btnAbout.Name = "btnAbout";
			this.btnAbout.Size = new System.Drawing.Size(64, 24);
			this.btnAbout.TabIndex = 11;
			this.btnAbout.Text = "about...";
			this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
			// 
			// txtItemValue
			// 
			this.txtItemValue.Location = new System.Drawing.Point(56, 296);
			this.txtItemValue.Name = "txtItemValue";
			this.txtItemValue.ReadOnly = true;
			this.txtItemValue.Size = new System.Drawing.Size(232, 20);
			this.txtItemValue.TabIndex = 7;
			this.txtItemValue.Text = "";
			// 
			// txtItemWriteRes
			// 
			this.txtItemWriteRes.Location = new System.Drawing.Point(368, 328);
			this.txtItemWriteRes.Name = "txtItemWriteRes";
			this.txtItemWriteRes.ReadOnly = true;
			this.txtItemWriteRes.Size = new System.Drawing.Size(152, 20);
			this.txtItemWriteRes.TabIndex = 8;
			this.txtItemWriteRes.Text = "";
			// 
			// opctreeicons
			// 
			this.opctreeicons.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit;
			this.opctreeicons.ImageSize = new System.Drawing.Size(16, 16);
			this.opctreeicons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("opctreeicons.ImageStream")));
			this.opctreeicons.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtItemTimeSt
			// 
			this.txtItemTimeSt.Location = new System.Drawing.Point(336, 272);
			this.txtItemTimeSt.Name = "txtItemTimeSt";
			this.txtItemTimeSt.ReadOnly = true;
			this.txtItemTimeSt.Size = new System.Drawing.Size(184, 20);
			this.txtItemTimeSt.TabIndex = 6;
			this.txtItemTimeSt.Text = "";
			// 
			// statusBarMain
			// 
			this.statusBarMain.Location = new System.Drawing.Point(0, 399);
			this.statusBarMain.Name = "statusBarMain";
			this.statusBarMain.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
																							 this.sbpTimeStart,
																							 this.sbpStatus,
																							 this.sbpDummy});
			this.statusBarMain.ShowPanels = true;
			this.statusBarMain.Size = new System.Drawing.Size(542, 16);
			this.statusBarMain.SizingGrip = false;
			this.statusBarMain.TabIndex = 9;
			// 
			// sbpTimeStart
			// 
			this.sbpTimeStart.MinWidth = 20;
			this.sbpTimeStart.Text = "0";
			this.sbpTimeStart.ToolTipText = "time of opc server start";
			this.sbpTimeStart.Width = 128;
			// 
			// btnItemMore
			// 
			this.btnItemMore.Enabled = false;
			this.btnItemMore.Location = new System.Drawing.Point(8, 360);
			this.btnItemMore.Name = "btnItemMore";
			this.btnItemMore.Size = new System.Drawing.Size(64, 24);
			this.btnItemMore.TabIndex = 11;
			this.btnItemMore.Text = "more...";
			this.btnItemMore.Click += new System.EventHandler(this.btnItemMore_Click);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 301);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 12;
			this.label4.Text = "value";
			this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(296, 276);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 16);
			this.label1.TabIndex = 12;
			this.label1.Text = "time";
			this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(296, 301);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 16);
			this.label2.TabIndex = 12;
			this.label2.Text = "type";
			this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// treeOpcItems
			// 
			this.treeOpcItems.HideSelection = false;
			this.treeOpcItems.ImageList = this.opctreeicons;
			this.treeOpcItems.Location = new System.Drawing.Point(8, 32);
			this.treeOpcItems.Name = "treeOpcItems";
			this.treeOpcItems.SelectedImageIndex = 1;
			this.treeOpcItems.Size = new System.Drawing.Size(248, 200);
			this.treeOpcItems.TabIndex = 2;
			this.treeOpcItems.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeOpcItems_AfterSelect);
			// 
			// txtItemSendValue
			// 
			this.txtItemSendValue.Location = new System.Drawing.Point(8, 328);
			this.txtItemSendValue.Name = "txtItemSendValue";
			this.txtItemSendValue.Size = new System.Drawing.Size(280, 20);
			this.txtItemSendValue.TabIndex = 10;
			this.txtItemSendValue.Text = "";
			// 
			// txtItemQual
			// 
			this.txtItemQual.Location = new System.Drawing.Point(56, 272);
			this.txtItemQual.Name = "txtItemQual";
			this.txtItemQual.ReadOnly = true;
			this.txtItemQual.Size = new System.Drawing.Size(232, 20);
			this.txtItemQual.TabIndex = 5;
			this.txtItemQual.Text = "";
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(542, 415);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.txtItemWriteRes,
																		  this.label3,
																		  this.txtItemSendValue,
																		  this.label1,
																		  this.treeOpcItems,
																		  this.txtServer,
																		  this.btnItemWrite,
																		  this.txtItemQual,
																		  this.txtItemID,
																		  this.txtItemValue,
																		  this.label2,
																		  this.txtItemDataType,
																		  this.txtItemTimeSt,
																		  this.listOpcView,
																		  this.txtServerInfo,
																		  this.statusBarMain,
																		  this.label4,
																		  this.btnItemMore,
																		  this.btnAbout});
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "MainForm";
			this.Text = "OPC.NET";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MainForm_Closing);
			this.Activated += new System.EventHandler(this.MainForm_Activated);
			((System.ComponentModel.ISupportInitialize)(this.sbpStatus)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpTimeStart)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpDummy)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

	/// <summary>
	/// The main entry point for the application.
	/// </summary>
	[STAThread]
	static void Main()
		{
		Application.Run( new MainForm() );
		}

	private void MainForm_Activated( object sender, System.EventArgs e )
		{
		if( ! first_activated )		// workaround to show SelServer form !after! MainForm is visible
			{
			first_activated = true;
			opc_connected	= DoInit();
			}
		}


	public TypeCode VT2TypeCode( VarEnum vevt )
		{
		switch( vevt )
			{
			case VarEnum.VT_I1:
				return TypeCode.SByte;
			case VarEnum.VT_I2:
				return TypeCode.Int16;
			case VarEnum.VT_I4:
				return TypeCode.Int32;
			case VarEnum.VT_I8:
				return TypeCode.Int64;

			case VarEnum.VT_UI1:
				return TypeCode.Byte;
			case VarEnum.VT_UI2:
				return TypeCode.UInt16;
			case VarEnum.VT_UI4:
				return TypeCode.UInt32;
			case VarEnum.VT_UI8:
				return TypeCode.UInt64;

			case VarEnum.VT_R4:
				return TypeCode.Single;
			case VarEnum.VT_R8:
				return TypeCode.Double;

			case VarEnum.VT_BSTR:
				return TypeCode.String;
			case VarEnum.VT_BOOL:
				return TypeCode.Boolean;
			case VarEnum.VT_DATE:
				return TypeCode.DateTime;
			case VarEnum.VT_DECIMAL:
				return TypeCode.Decimal;
			case VarEnum.VT_CY:				// not supported
				return TypeCode.Double;
			}

		return TypeCode.Object;
		}

	private void MainForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		if( ! opc_connected )
			return;
			
		if( theGrp != null )
			{
			theGrp.DataChanged -= new DataChangeEventHandler( this.theGrp_DataChange );
			theGrp.WriteCompleted -= new WriteCompleteEventHandler( this.theGrp_WriteComplete );
			RemoveItem();
			theGrp.Remove( false );
			theGrp = null;
			}

		if( theSrv != null )
			{
			theSrv.Disconnect();				// should clean up
			theSrv = null;
			}

		opc_connected	= false;
		}

	private void btnItemWrite_Click(object sender, System.EventArgs e)
		{
		try
			{
			txtItemWriteRes.Text = "";
			
			// convert the user text to OPC data type of item
			object[]	arrVal = new Object[1];
			arrVal[0] = Convert.ChangeType( txtItemSendValue.Text, itmTypeCode );

			int[]	serverhandles	= new int[1] { itmHandleServer };
			int		cancelID;
			int[]	arrErr;
			theGrp.Write( serverhandles, arrVal, 9988, out cancelID, out arrErr );

			GC.Collect();		// just for fun
			}
		catch( FormatException )
			{
			MessageBox.Show( this, "Invalid data format!", "opcItemDoWrite_Click", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
			}
		catch( OverflowException )
			{
			MessageBox.Show( this, "Invalid data range/overflow!", "opcItemDoWrite_Click", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
			}
		catch( COMException )
			{
			MessageBox.Show( this, "OPC Write Item error!", "opcItemDoWrite_Click", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
			}
		}

	private void listOpcView_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		if( listOpcView.SelectedItems.Count != 1 )		// no selection
			return;

		ListViewItem	selitem = listOpcView.SelectedItems[0];
		selecteditem = selitem.Text;				// new selected item in ListView
		itmFullID = selitem.SubItems[1].Text;		// fully qualified OPC item namespace path

		ViewItem( itmFullID );		// show item data+attribs on screen
		}

	private void treeOpcItems_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
		{
		listOpcView.Items.Clear();
		RemoveItem();				// remove item from group

		try
			{
			theSrv.ChangeBrowsePosition( OPCBROWSEDIRECTION.OPC_BROWSE_TO, "" );	// to root

			if( e.Node.FullPath.Length > rootname.Length )		// check if it's only the dummy root
				{
				selectednode = e.Node.FullPath.Substring( rootname.Length + 1 );
				string[] splitpath = selectednode.Split( new char[] {'\t'} );	// convert path-string to string-array (separator)

				foreach( string n in splitpath )
					theSrv.ChangeBrowsePosition( OPCBROWSEDIRECTION.OPC_BROWSE_DOWN, n );	// browse to node in OPC namespace
				}
			else
				selectednode = "";

			// get all items at this node level
			ArrayList lst;
			theSrv.Browse( OPCBROWSETYPE.OPC_LEAF, out lst );
			if( lst == null )
				return;
			if( lst.Count < 1 )
				return;
			
			// enum+add all item names to ListView
			string[]	itemstrings = new string[ 2 ];
			foreach( string item in lst )
				{
				itemstrings[0] = item;
				itemstrings[1] = theSrv.GetItemID( item );
				listOpcView.Items.Add( new ListViewItem( itemstrings, 0 ) );
				}

			// preselect top item in ListView
			listOpcView.Items[0].Selected = true;
			}
		catch( Exception ex )		// exceptions MUST be handled
			{
			MessageBox.Show( this, "browse error! " + ex.ToString(), "Exception browsing namespace", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
			}
		}

	private void btnItemMore_Click( object sender, System.EventArgs e )
		{
		PropsForm	frmProps = new PropsForm( ref theSrv, itmFullID );		// create item properties form
		frmProps.ShowDialog( this );
		}

	private void btnAbout_Click(object sender, System.EventArgs e)
		{
		AboutForm	frmAbout = new AboutForm();			// show about dialog
		frmAbout.ShowDialog( this );
		}

}	// class MainForm

}	// namespace DirectOPCClient
