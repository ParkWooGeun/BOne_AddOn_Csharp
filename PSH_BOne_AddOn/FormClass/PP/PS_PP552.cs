using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 공정상태 관리
	/// </summary>
	internal class PS_PP552 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_PP552L;
		private SAPbouiCOM.DBDataSource oDS_PS_PP552M;
		private SAPbouiCOM.DBDataSource oDS_PS_PP552N;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{

			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP552.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP552_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP552");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP552_CreateItems();
				PS_PP552_SetComboBox();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Items.Item("Folder01").Click();
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP552_CreateItems
		/// </summary>
		private void PS_PP552_CreateItems()
		{
			try
			{
				oDS_PS_PP552L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP552M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
				oDS_PS_PP552N = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");

				// 메트릭스 개체 할당
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				oMat03 = oForm.Items.Item("Mat03").Specific;
				oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat03.AutoResizeColumns();

				//공정상태 삭제
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode01").Specific.DataBind.SetBound(true, "", "CpCode01");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName01").Specific.DataBind.SetBound(true, "", "CpName01");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode01").Specific.DataBind.SetBound(true, "", "CntcCode01");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcName01").Specific.DataBind.SetBound(true, "", "CntcName01");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkGbn01").Specific.DataBind.SetBound(true, "", "WorkGbn01");

				//공정상태
				oForm.DataSources.UserDataSources.Add("CpStatus01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpStatus01").Specific.DataBind.SetBound(true, "", "CpStatus01");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType01").Specific.DataBind.SetBound(true, "", "CardType01");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType01").Specific.DataBind.SetBound(true, "", "ItemType01");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum01").Specific.DataBind.SetBound(true, "", "OrdNum01");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("OrdSub101", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("OrdSub101").Specific.DataBind.SetBound(true, "", "OrdSub101");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("OrdSub201", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("OrdSub201").Specific.DataBind.SetBound(true, "", "OrdSub201");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName01").Specific.DataBind.SetBound(true, "", "ItemName01");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec01").Specific.DataBind.SetBound(true, "", "ItemSpec01");

				//공정상태 비고 등록
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode02").Specific.DataBind.SetBound(true, "", "CpCode02");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName02").Specific.DataBind.SetBound(true, "", "CpName02");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode02").Specific.DataBind.SetBound(true, "", "CntcCode02");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcName02").Specific.DataBind.SetBound(true, "", "CntcName02");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkGbn02").Specific.DataBind.SetBound(true, "", "WorkGbn02");

				//공정상태
				oForm.DataSources.UserDataSources.Add("CpStatus02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpStatus02").Specific.DataBind.SetBound(true, "", "CpStatus02");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType02").Specific.DataBind.SetBound(true, "", "CardType02");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType02").Specific.DataBind.SetBound(true, "", "ItemType02");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum02").Specific.DataBind.SetBound(true, "", "OrdNum02");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("OrdSub102", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("OrdSub102").Specific.DataBind.SetBound(true, "", "OrdSub102");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("OrdSub202", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("OrdSub202").Specific.DataBind.SetBound(true, "", "OrdSub202");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName02").Specific.DataBind.SetBound(true, "", "ItemName02");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec02").Specific.DataBind.SetBound(true, "", "ItemSpec02");

				//공정상태 계획공수 등록
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID03").Specific.DataBind.SetBound(true, "", "BPLID03");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt03").Specific.DataBind.SetBound(true, "", "FrDt03");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt03").Specific.DataBind.SetBound(true, "", "ToDt03");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode03").Specific.DataBind.SetBound(true, "", "CpCode03");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName03").Specific.DataBind.SetBound(true, "", "CpName03");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode03").Specific.DataBind.SetBound(true, "", "CntcCode03");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcName03").Specific.DataBind.SetBound(true, "", "CntcName03");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkGbn03").Specific.DataBind.SetBound(true, "", "WorkGbn03");

				//공정상태
				oForm.DataSources.UserDataSources.Add("CpStatus03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpStatus03").Specific.DataBind.SetBound(true, "", "CpStatus03");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType03").Specific.DataBind.SetBound(true, "", "CardType03");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType03").Specific.DataBind.SetBound(true, "", "ItemType03");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum03").Specific.DataBind.SetBound(true, "", "OrdNum03");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("OrdSub103", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("OrdSub103").Specific.DataBind.SetBound(true, "", "OrdSub103");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("OrdSub203", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("OrdSub203").Specific.DataBind.SetBound(true, "", "OrdSub203");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName03").Specific.DataBind.SetBound(true, "", "ItemName03");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec03").Specific.DataBind.SetBound(true, "", "ItemSpec03");
				
				//일자기본SET
				oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.AddMonths(-2).ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.Value = DateTime.Now.AddMonths(-2).ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt03").Specific.Value = DateTime.Now.AddMonths(-2).ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt03").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("CpStatus03").Enabled = false;	//계획공수 등록 시는 공정상태가 대기인 자료만 조회하기 위한 콤보 선택 불가 조치
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP552_SetComboBox
		/// </summary>
		private void PS_PP552_SetComboBox()
		{
			string sQry;
			string BPLID;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//공정상태 삭제
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//작업구분
				sQry = " SELECT     Code AS [Code], ";
				sQry += "           Name AS [Name]";
				sQry += " FROM      [@PSH_ITMBSORT]";
				sQry += " WHERE     U_PudYN = 'Y'";
				oForm.Items.Item("WorkGbn01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn01").Specific, sQry, "%", false, false);

				//공정상태
				oForm.Items.Item("CpStatus01").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("CpStatus01").Specific.ValidValues.Add("10", "대기");
				oForm.Items.Item("CpStatus01").Specific.ValidValues.Add("20", "시작");
				oForm.Items.Item("CpStatus01").Specific.ValidValues.Add("30", "완료");
				oForm.Items.Item("CpStatus01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래처구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				oForm.Items.Item("CardType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType01").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				oForm.Items.Item("ItemType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType01").Specific, sQry, "%", false, false);

				//매트릭스
				//공정상태
				oMat01.Columns.Item("CpStatus").ValidValues.Add("10", "대기");
				oMat01.Columns.Item("CpStatus").ValidValues.Add("20", "시작");
				oMat01.Columns.Item("CpStatus").ValidValues.Add("30", "완료");

				//공정상태 비고 등록
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//작업구분
				sQry = "  SELECT     Code AS [Code], ";
				sQry += "            Name AS [Name]";
				sQry += " FROM      [@PSH_ITMBSORT]";
				sQry += " WHERE     U_PudYN = 'Y'";
				oForm.Items.Item("WorkGbn02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn02").Specific, sQry, "%", false, false);

				//공정상태
				oForm.Items.Item("CpStatus02").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("CpStatus02").Specific.ValidValues.Add("10", "대기");
				oForm.Items.Item("CpStatus02").Specific.ValidValues.Add("20", "시작");
				oForm.Items.Item("CpStatus02").Specific.ValidValues.Add("30", "완료");
				oForm.Items.Item("CpStatus02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래처구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				oForm.Items.Item("CardType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType02").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				oForm.Items.Item("ItemType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType02").Specific, sQry, "%", false, false);

				//매트릭스
				//공정상태
				oMat02.Columns.Item("CpStatus").ValidValues.Add("10", "대기");
				oMat02.Columns.Item("CpStatus").ValidValues.Add("20", "시작");
				oMat02.Columns.Item("CpStatus").ValidValues.Add("30", "완료");

				//공정상태 계획공수 등록
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID03").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//작업구분
				sQry = " SELECT     Code AS [Code], ";
				sQry += "           Name AS [Name]";
				sQry += " FROM      [@PSH_ITMBSORT]";
				sQry += " WHERE     U_PudYN = 'Y'";
				oForm.Items.Item("WorkGbn03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn03").Specific, sQry, "%", false, false);

				//공정상태
				oForm.Items.Item("CpStatus03").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("CpStatus03").Specific.ValidValues.Add("10", "대기");
				oForm.Items.Item("CpStatus03").Specific.ValidValues.Add("20", "시작");
				oForm.Items.Item("CpStatus03").Specific.ValidValues.Add("30", "완료");
				oForm.Items.Item("CpStatus03").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래처구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				oForm.Items.Item("CardType03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType03").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "           U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				oForm.Items.Item("ItemType03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType03").Specific, sQry, "%", false, false);

				//매트릭스
				//공정상태
				oMat03.Columns.Item("CpStatus").ValidValues.Add("10", "대기");
				oMat03.Columns.Item("CpStatus").ValidValues.Add("20", "시작");
				oMat03.Columns.Item("CpStatus").ValidValues.Add("30", "완료");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP552_SaveData02  비고 등록
		/// </summary>
		/// <returns></returns>
		private bool PS_PP552_SaveData02()
		{
			bool returnValue = false;
			short loopCount;
			string sQry;
			string MainOrdNum;	//작번
			string SubOrdNum01;	//서브작번1
			string SubOrdNum02;	//서브작번2
			string CpCode;		//공정코드
			string CpCount;		//공정횟수
			string CpStatus;    //공정상태
			string RegDate;     //등록일
			string Comment;		//비고

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				ProgressBar01.Text = "저장중...";

				oMat02.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat02.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP552M.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						MainOrdNum  = oDS_PS_PP552M.GetValue("U_ColReg02", loopCount).ToString().Trim(); //작번
						SubOrdNum01 = oDS_PS_PP552M.GetValue("U_ColReg03", loopCount).ToString().Trim(); //서브작번1
						SubOrdNum02 = oDS_PS_PP552M.GetValue("U_ColReg04", loopCount).ToString().Trim(); //서브작번2
						CpCode      = oDS_PS_PP552M.GetValue("U_ColReg08", loopCount).ToString().Trim(); //공정코드
						CpCount     = Convert.ToString(Convert.ToDouble(oDS_PS_PP552M.GetValue("U_ColReg10", loopCount).ToString().Trim()) - 1); //공정횟수
						CpStatus    = oDS_PS_PP552M.GetValue("U_ColReg11", loopCount).ToString().Trim(); //공정상태
						RegDate     = oDS_PS_PP552M.GetValue("U_ColDt01", loopCount).ToString().Trim(); //등록일
						Comment     = oDS_PS_PP552M.GetValue("U_ColReg12", loopCount).ToString().Trim(); //비고

						sQry = "EXEC [PS_PP552_12] '";
						sQry += MainOrdNum + "','";
						sQry += SubOrdNum01 + "','";
						sQry += SubOrdNum02 + "','";
						sQry += CpCode + "','";
						sQry += CpCount + "','";
						sQry += CpStatus + "','";
						sQry += RegDate + "','";
						sQry += Comment + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.MessageBox("등록 완료!");
				returnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP552_SaveData03  계획공수 등록
		/// </summary>
		/// <returns></returns>
		private bool PS_PP552_SaveData03()
		{
			bool returnValue = false;
			short loopCount;
			string sQry;
			string MainOrdNum;  //작번
			string SubOrdNum01; //서브작번1
			string SubOrdNum02; //서브작번2
			string CpCode;      //공정코드
			string CpCount;     //공정횟수
			string CpStatus;    //공정상태
			double PlanHour;    //계획공수

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				ProgressBar01.Text = "저장중...";

				oMat03.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat03.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP552N.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						MainOrdNum  = oDS_PS_PP552N.GetValue("U_ColReg02", loopCount).ToString().Trim();					//작번
						SubOrdNum01 = oDS_PS_PP552N.GetValue("U_ColReg03", loopCount).ToString().Trim();					//서브작번1
						SubOrdNum02 = oDS_PS_PP552N.GetValue("U_ColReg04", loopCount).ToString().Trim();					//서브작번2
						CpCode      = oDS_PS_PP552N.GetValue("U_ColReg08", loopCount).ToString().Trim();					//공정코드
						CpCount     = Convert.ToString(Convert.ToDouble(oDS_PS_PP552N.GetValue("U_ColReg10", loopCount).ToString().Trim()) - 1); //공정횟수
						CpStatus    = oDS_PS_PP552N.GetValue("U_ColReg11", loopCount).ToString().Trim();					//공정상태
						PlanHour    = Convert.ToDouble(oDS_PS_PP552N.GetValue("U_ColQty01", loopCount).ToString().Trim());  //계획공수

						sQry = "EXEC [PS_PP552_13] '";
						sQry += MainOrdNum + "','";
						sQry += SubOrdNum01 + "','";
						sQry += SubOrdNum02 + "','";
						sQry += CpCode + "','";
						sQry += CpCount + "','";
						sQry += CpStatus + "','";
						sQry += PlanHour + "'";

						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.MessageBox("등록 완료!");
				returnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP552_DeleteData01  공정정보 삭제
		/// </summary>
		/// <returns></returns>
		private bool PS_PP552_DeleteData01()
		{
			bool returnValue = false;
			short loopCount;
			string sQry;
			string errMessage = string.Empty;
			string MainOrdNum;  //작번
			string SubOrdNum01; //서브작번1
			string SubOrdNum02; //서브작번2
			string CpCode;      //공정코드
			string CpCount;     //공정횟수
			string CpStatus;    //공정상태
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oMat01.VisualRowCount == 0)
				{
					errMessage = "삭제대상이 없습니다. 확인하세요.";
					throw new Exception();
				}

				oMat01.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{

					if (oDS_PS_PP552L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						MainOrdNum  = oDS_PS_PP552L.GetValue("U_ColReg02", loopCount).ToString().Trim();
						SubOrdNum01 = oDS_PS_PP552L.GetValue("U_ColReg03", loopCount).ToString().Trim();
						SubOrdNum02 = oDS_PS_PP552L.GetValue("U_ColReg04", loopCount).ToString().Trim();
						CpCode      = oDS_PS_PP552L.GetValue("U_ColReg08", loopCount).ToString().Trim();
						CpCount     = Convert.ToString(Convert.ToDouble(oDS_PS_PP552L.GetValue("U_ColReg10", loopCount).ToString().Trim()) - 1);
						CpStatus    = oDS_PS_PP552L.GetValue("U_ColReg11", loopCount).ToString().Trim();

						sQry = "EXEC [PS_PP552_31] ";
						sQry += "'" + MainOrdNum + "',";
						sQry += "'" + SubOrdNum01 + "',";
						sQry += "'" + SubOrdNum02 + "',";
						sQry += "'" + CpCode + "',";
						sQry += "'" + CpCount + "',";
						sQry += "'" + CpStatus + "'";

						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
				returnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Update();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP552_CheckAfterStatus 이후 공정상태가 존재하는지 확인
		/// </summary>
		/// <param name="pRow"></param>
		/// <returns></returns>
		private bool PS_PP552_CheckAfterStatus(int pRow)
		{
			bool returnValue = false;
			string sQry;
			string MainOrdNum;  //작번
			string SubOrdNum01; //서브작번1
			string SubOrdNum02; //서브작번2
			string CpCode;      //공정코드
			string CpCount;     //공정횟수
			string CpStatus;    //공정상태
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				MainOrdNum  = oDS_PS_PP552L.GetValue("U_ColReg02", pRow - 1).ToString().Trim();
				SubOrdNum01 = oDS_PS_PP552L.GetValue("U_ColReg03", pRow - 1).ToString().Trim();
				SubOrdNum02 = oDS_PS_PP552L.GetValue("U_ColReg04", pRow - 1).ToString().Trim();
				CpCode      = oDS_PS_PP552L.GetValue("U_ColReg08", pRow - 1).ToString().Trim();
				CpCount     = Convert.ToString(Convert.ToDouble(oDS_PS_PP552L.GetValue("U_ColReg10", pRow - 1).ToString().Trim()) - 1);
				CpStatus    = oDS_PS_PP552L.GetValue("U_ColReg11", pRow - 1).ToString().Trim();

				sQry = "EXEC [PS_PP552_91] '";
				sQry += MainOrdNum + "','";
				sQry += SubOrdNum01 + "','";
				sQry += SubOrdNum02 + "','";
				sQry += CpCode + "','";
				sQry += CpCount + "','";
				sQry += CpStatus + "'";

				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item("AfterStatusYN").Value == "Y")
				{
					returnValue = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP552_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP552_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			string OrdNum;
			string OrdSub1;
			string OrdSub2;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oUID == "Mat01")
				{
					oMat01.AutoResizeColumns();
				}
				else if (oUID == "Mat02")
				{
					oMat02.AutoResizeColumns();
				}
				else if (oUID == "Mat03")
				{
					oMat03.AutoResizeColumns();
				}
				else if (oUID == "CntcCode01")
				{
					oForm.Items.Item("CntcName01").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CntcCode02")
				{
					oForm.Items.Item("CntcName02").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CntcCode03")
				{
					oForm.Items.Item("CntcName03").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode01")
				{
					oForm.Items.Item("CpName01").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode02")
				{
					oForm.Items.Item("CpName02").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode03")
				{
					oForm.Items.Item("CpName03").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "OrdNum01" || oUID == "OrdSub101" || oUID == "OrdSub201")
				{
					OrdNum = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();
					OrdSub1 = oForm.Items.Item("OrdSub101").Specific.Value.ToString().Trim();
					OrdSub2 = oForm.Items.Item("OrdSub201").Specific.Value.ToString().Trim();

					sQry = " SELECT   CASE";
					sQry += "            WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "            ELSE T0.U_JakMyung";
					sQry += "        END AS [ItemName],";
					sQry += "        CASE";
					sQry += "            WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "            ELSE T0.U_JakSize";
					sQry += "        END AS [ItemSpec]";
					sQry += " FROM     [@PS_PP020H] AS T0";
					sQry += " WHERE   T0.U_JakName = '" + OrdNum + "'";
					sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + OrdSub1 + "' = '' THEN '00' ELSE '" + OrdSub1 + "' END";
					sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + OrdSub2 + "' = '' THEN '000' ELSE '" + OrdSub2 + "' END";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName01").Specific.Value = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
					oForm.Items.Item("ItemSpec01").Specific.Value = oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim();
				}
				else if (oUID == "OrdNum02" || oUID == "OrdSub102" || oUID == "OrdSub202")
				{
					OrdNum  = oForm.Items.Item("OrdNum02").Specific.Value.ToString().Trim();
					OrdSub1 = oForm.Items.Item("OrdSub102").Specific.Value.ToString().Trim();
					OrdSub2 = oForm.Items.Item("OrdSub202").Specific.Value.ToString().Trim();

					sQry = " SELECT   CASE";
					sQry += "           WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "           ELSE T0.U_JakMyung";
					sQry += "       END AS [ItemName],";
					sQry += "       CASE";
					sQry += "           WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "           ELSE T0.U_JakSize";
					sQry += "       END AS [ItemSpec]";
					sQry += " FROM     [@PS_PP020H] AS T0";
					sQry += " WHERE   T0.U_JakName = '" + OrdNum + "'";
					sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + OrdSub1 + "' = '' THEN '00' ELSE '" + OrdSub1 + "' END";
					sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + OrdSub2 + "' = '' THEN '000' ELSE '" + OrdSub2 + "' END";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName02").Specific.Value = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
					oForm.Items.Item("ItemSpec02").Specific.Value = oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim();
				}
				else if (oUID == "OrdNum03" || oUID == "OrdSub103" || oUID == "OrdSub203")
				{
					OrdNum =  oForm.Items.Item("OrdNum03").Specific.Value.ToString().Trim();
					OrdSub1 = oForm.Items.Item("OrdSub103").Specific.Value.ToString().Trim();
					OrdSub2 = oForm.Items.Item("OrdSub203").Specific.Value.ToString().Trim();

					sQry = " SELECT   CASE";
					sQry += "           WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "           ELSE T0.U_JakMyung";
					sQry += "       END AS [ItemName],";
					sQry += "       CASE";
					sQry += "           WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "           ELSE T0.U_JakSize";
					sQry += "       END AS [ItemSpec]";
					sQry += " FROM     [@PS_PP020H] AS T0";
					sQry += " WHERE   T0.U_JakName = '" + OrdNum + "'";
					sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + OrdSub1 + "' = '' THEN '00' ELSE '" + OrdSub1 + "' END";
					sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + OrdSub2 + "' = '' THEN '000' ELSE '" + OrdSub2 + "' END";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName03").Specific.Value = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
					oForm.Items.Item("ItemSpec03").Specific.Value = oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP552_ResizeForm
		/// </summary>
		private void PS_PP552_ResizeForm()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Mat01").Height + 170;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Mat01").Width + 30;

				oMat01.AutoResizeColumns();
				oMat02.AutoResizeColumns();
				oMat03.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP552_MTX01 공정상태 삭제 대상 조회
		/// </summary>
		private void PS_PP552_MTX01()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string BPLID;		 //사업장
			string FrDt;		 //기간(시작)
			string ToDt;		 //기간(종료)
			string CpCode;		 //공정
			string WorkerCode;	 //등록자
			string WorkGbn;		 //작업구분
			string CpStatus;	 //공정상태
			string CardType;	 //거래처구분
			string ItemType;	 //품목구분
			string MainOrdNum;	 //메인작번
			string SubOrdNum01;  //서브작번1
			string SubOrdNum02;  //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				BPLID       = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				FrDt        = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt        = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				CpCode      = oForm.Items.Item("CpCode01").Specific.Value.ToString().Trim();
				WorkerCode  = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				WorkGbn     = oForm.Items.Item("WorkGbn01").Specific.Value.ToString().Trim();
				CpStatus    = oForm.Items.Item("CpStatus01").Specific.Value.ToString().Trim();
				CardType    = oForm.Items.Item("CardType01").Specific.Value.ToString().Trim();
				ItemType    = oForm.Items.Item("ItemType01").Specific.Value.ToString().Trim();
				MainOrdNum  = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();
				SubOrdNum01 = oForm.Items.Item("OrdSub101").Specific.Value.ToString().Trim();
				SubOrdNum02 = oForm.Items.Item("OrdSub201").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP552_01] '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += WorkerCode + "','";
				sQry += WorkGbn + "','";
				sQry += CpStatus + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += MainOrdNum + "','";
				sQry += SubOrdNum01 + "','";
				sQry += SubOrdNum02 + "'";

				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_PP552L.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount + 1 > oDS_PS_PP552L.Size)
					{
						oDS_PS_PP552L.InsertRecord(loopCount);
					}

					oMat01.AddRow();
					oDS_PS_PP552L.Offset = loopCount;
					oDS_PS_PP552L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
					oDS_PS_PP552L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Check").Value.ToString().Trim());	  //선택
					oDS_PS_PP552L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	  //작번
					oDS_PS_PP552L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());	  //서브작번1
					oDS_PS_PP552L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());	  //서브작번2
					oDS_PS_PP552L.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("WkOrdQty").Value.ToString().Trim());  //작업지시수량
					oDS_PS_PP552L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());	  //작명
					oDS_PS_PP552L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());  //품명
					oDS_PS_PP552L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());  //규격
					oDS_PS_PP552L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	  //공정코드
					oDS_PS_PP552L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	  //공정명
					oDS_PS_PP552L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("CpCount").Value.ToString().Trim());	  //공정횟수
					oDS_PS_PP552L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("CpStatus").Value.ToString().Trim());  //공정상태
					oDS_PS_PP552L.SetValue("U_ColDt01", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("RegDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일
					oDS_PS_PP552L.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("WorkerCode").Value.ToString().Trim());  //등록자사번
					oDS_PS_PP552L.SetValue("U_ColReg13", loopCount, oRecordSet.Fields.Item("WorkerName").Value.ToString().Trim());  //등록자성명

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP552_MTX02 공정상태 비고 등록 대상 조회
		/// </summary>
		private void PS_PP552_MTX02()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string BPLID;        //사업장
			string FrDt;         //기간(시작)
			string ToDt;         //기간(종료)
			string CpCode;       //공정
			string WorkerCode;   //등록자
			string WorkGbn;      //작업구분
			string CpStatus;     //공정상태
			string CardType;     //거래처구분
			string ItemType;     //품목구분
			string MainOrdNum;   //메인작번
			string SubOrdNum01;  //서브작번1
			string SubOrdNum02;  //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID       = oForm.Items.Item("BPLID02").Specific.Selected.Value.ToString().Trim();
				FrDt        = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt        = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				CpCode      = oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim();
				WorkerCode  = oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim();
				WorkGbn     = oForm.Items.Item("WorkGbn02").Specific.Value.ToString().Trim();
				CpStatus    = oForm.Items.Item("CpStatus02").Specific.Value.ToString().Trim();
				CardType    = oForm.Items.Item("CardType02").Specific.Value.ToString().Trim();
				ItemType    = oForm.Items.Item("ItemType02").Specific.Value.ToString().Trim();
				MainOrdNum  = oForm.Items.Item("OrdNum02").Specific.Value.ToString().Trim();
				SubOrdNum01 = oForm.Items.Item("OrdSub102").Specific.Value.ToString().Trim();
				SubOrdNum02 = oForm.Items.Item("OrdSub202").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "EXEC [PS_PP552_02] '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += WorkerCode + "','";
				sQry += WorkGbn + "','";
				sQry += CpStatus + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += MainOrdNum + "','";
				sQry += SubOrdNum01 + "','";
				sQry += SubOrdNum02 + "'";

				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oDS_PS_PP552M.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount + 1 > oDS_PS_PP552M.Size)
					{
						oDS_PS_PP552M.InsertRecord(loopCount);
					}

					oMat02.AddRow();
					oDS_PS_PP552M.Offset = loopCount;
					oDS_PS_PP552M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
					oDS_PS_PP552M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Check").Value.ToString().Trim());	 //선택
					oDS_PS_PP552M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP552M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());	 //서브작번1
					oDS_PS_PP552M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());	 //서브작번2
					oDS_PS_PP552M.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("WkOrdQty").Value.ToString().Trim()); //작업지시수량
					oDS_PS_PP552M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());	 //작명
					oDS_PS_PP552M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP552M.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim()); //규격
					oDS_PS_PP552M.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	 //공정코드
					oDS_PS_PP552M.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	 //공정명
					oDS_PS_PP552M.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("CpCount").Value.ToString().Trim());	 //공정횟수
					oDS_PS_PP552M.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("CpStatus").Value.ToString().Trim()); //공정상태
				    oDS_PS_PP552M.SetValue("U_ColDt01", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("RegDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일
					oDS_PS_PP552M.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());	 //비고

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP552_MTX03 공정상태 계획공수 등록 대상 조회
		/// </summary>
		private void PS_PP552_MTX03()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string BPLID;        //사업장
			string FrDt;         //기간(시작)
			string ToDt;         //기간(종료)
			string CpCode;       //공정
			string WorkerCode;   //등록자
			string WorkGbn;      //작업구분
			string CpStatus;     //공정상태
			string CardType;     //거래처구분
			string ItemType;     //품목구분
			string MainOrdNum;   //메인작번
			string SubOrdNum01;  //서브작번1
			string SubOrdNum02;  //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID03").Specific.Selected.Value.ToString().Trim();
				FrDt        = oForm.Items.Item("FrDt03").Specific.Value.ToString().Trim();
				ToDt        = oForm.Items.Item("ToDt03").Specific.Value.ToString().Trim();
				CpCode      = oForm.Items.Item("CpCode03").Specific.Value.ToString().Trim();
				WorkerCode  = oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim();
				WorkGbn     = oForm.Items.Item("WorkGbn03").Specific.Value.ToString().Trim();
				CpStatus    = oForm.Items.Item("CpStatus03").Specific.Value.ToString().Trim();
				CardType    = oForm.Items.Item("CardType03").Specific.Value.ToString().Trim();
				ItemType    = oForm.Items.Item("ItemType03").Specific.Value.ToString().Trim();
				MainOrdNum  = oForm.Items.Item("OrdNum03").Specific.Value.ToString().Trim();
				SubOrdNum01 = oForm.Items.Item("OrdSub103").Specific.Value.ToString().Trim();
				SubOrdNum02 = oForm.Items.Item("OrdSub203").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP552_03] '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += WorkerCode + "','";
				sQry += WorkGbn + "','";
				sQry += CpStatus + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += MainOrdNum + "','";
				sQry += SubOrdNum01 + "','";
				sQry += SubOrdNum02 + "'";

				oRecordSet.DoQuery(sQry);

				oMat03.Clear();
				oDS_PS_PP552N.Clear();
				oMat03.FlushToDataSource();
				oMat03.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount + 1 > oDS_PS_PP552N.Size)
					{
						oDS_PS_PP552N.InsertRecord(loopCount);
					}

					oMat03.AddRow();
					oDS_PS_PP552N.Offset = loopCount;
					oDS_PS_PP552N.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
					oDS_PS_PP552N.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Check").Value.ToString().Trim());	 //선택
					oDS_PS_PP552N.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP552N.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());	 //서브작번1
					oDS_PS_PP552N.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());	 //서브작번2
					oDS_PS_PP552N.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("WkOrdQty").Value.ToString().Trim()); //작업지시수량
					oDS_PS_PP552N.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());	 //작명
					oDS_PS_PP552N.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP552N.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim()); //규격
					oDS_PS_PP552N.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	 //공정코드
					oDS_PS_PP552N.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	 //공정명
					oDS_PS_PP552N.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("CpCount").Value.ToString().Trim());	 //공정횟수
					oDS_PS_PP552N.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("CpStatus").Value.ToString().Trim()); //공정상태
				    oDS_PS_PP552N.SetValue("U_ColDt01", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("RegDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일
					oDS_PS_PP552N.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("PlanHour").Value.ToString().Trim()); //계획공수

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat03.LoadFromDataSource();
				oMat03.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Form Item Event
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">pVal</param>
		/// <param name="BubbleEvent">Bubble Event</param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			switch (pVal.EventType)
			{
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
            }
		}

		/// <summary>
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSrch01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP552_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnDel01")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							if (PS_PP552_DeleteData01() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_PP552_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP552_MTX02();
						}
					}
					else if (pVal.ItemUID == "BtnSave02")
					{
						if (PS_PP552_SaveData02() == false)
						{
							BubbleEvent = false;
							return;
						}
						PS_PP552_MTX02();
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP552_MTX03();
						}
					}
					else if (pVal.ItemUID == "BtnSave03")
					{
						if (PS_PP552_SaveData03() == false)
						{
							BubbleEvent = false;
							return;
						}
						PS_PP552_MTX03();
					}
					else if (pVal.ItemUID == "Mat01" && pVal.ColUID == "Check" && pVal.Row > 0)
					{
						if (oMat01.RowCount >= pVal.Row)  //빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
						{
							if (PS_PP552_CheckAfterStatus(pVal.Row) == true)
							{
								PSH_Globals.SBO_Application.MessageBox("선택한 작번은 이후 공정상태가 존재합니다. 선택할 수 없습니다.");
								oMat01.Columns.Item("Check").Cells.Item(pVal.Row).Specific.Checked = false;
								return;
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum01", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum02", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode03", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode03", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum03", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat01.SelectRow(pVal.Row, true, false);
						}
					}
					else if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row > 0)
						{
							oMat02.SelectRow(pVal.Row, true, false);
						}
					}
					else if (pVal.ItemUID == "Mat03")
					{
						if (pVal.Row > 0)
						{
							oMat03.SelectRow(pVal.Row, true, false);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							PS_PP552_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_PP552_FlushToItemValue(pVal.ItemUID, 0, "");
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP552_ResizeForm();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP552L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP552M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP552N);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
