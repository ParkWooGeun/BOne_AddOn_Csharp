using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산일보조회
	/// </summary>
	internal class PS_PP045 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Item_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Col_Row;	
		private int oLast_Mode;
		private int oSeq;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP045.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP045_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP045");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

                PS_PP045_CreateItems();
                PS_PP045_SetComboBox();

                oForm.EnableMenu("1281", false); // 찾기
				oForm.EnableMenu("1282", false); // 추가
			}
			catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP045_CreateItems()
        {
            try
            {
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01"); //디비데이터 소스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific; //메트릭스 개체 할당

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //팀
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                //반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                //작업타입
                oForm.DataSources.UserDataSources.Add("OrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("OrdGbn").Specific.DataBind.SetBound(true, "", "OrdGbn");

                //일자(시작)
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM01");

                //일자(종료)
                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");

                //작업자(사번)
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //작업자(성명)
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP045_SetComboBox()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //작업구분
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code";
                oRecordSet01.DoQuery(sQry);
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("", "");
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("OrdGbn").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("OrdGbn").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //작업구분(matrix)
                oMat01.Columns.Item("OrdType").ValidValues.Add("10", "일반");
                oMat01.Columns.Item("OrdType").ValidValues.Add("20", "PSMT지원");
                oMat01.Columns.Item("OrdType").ValidValues.Add("30", "외주");
                oMat01.Columns.Item("OrdType").ValidValues.Add("40", "실적");
                oMat01.Columns.Item("OrdType").ValidValues.Add("50", "일반조정");
                oMat01.Columns.Item("OrdType").ValidValues.Add("60", "외주조정");
                oMat01.Columns.Item("OrdType").ValidValues.Add("70", "설계시간");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP045_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i;
            string sQry;
            string BPLId;
            string TeamCode;
            string RspCode;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "BPLId":

                        BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //부서콤보세팅
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT    U_Code AS [Code],";
                        sQry += "           U_CodeNm As [Name]";
                        sQry += " FROM      [@PS_HR200L]";
                        sQry += " WHERE     Code = '1'";
                        sQry += "           AND U_UseYN = 'Y'";
                        sQry += "           AND U_Char2 = '" + BPLId + "'";
                        sQry += " ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "TeamCode":

                        TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //담당콤보세팅
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT    U_Code AS [Code],";
                        sQry += "           U_CodeNm As [Name]";
                        sQry += " FROM      [@PS_HR200L]";
                        sQry += " WHERE     Code = '2'";
                        sQry += "           AND U_UseYN = 'Y'";
                        sQry += "           AND U_Char1 = '" + TeamCode + "'";
                        sQry += " ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "RspCode":

                        TeamCode = oForm.Items.Item("TeamCode").Specific.Value;
                        RspCode = oForm.Items.Item("RspCode").Specific.Value;

                        if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //반콤보세팅
                        oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT    U_Code AS [Code],";
                        sQry += "           U_CodeNm As [Name]";
                        sQry += " FROM      [@PS_HR200L]";
                        sQry += " WHERE     Code = '9'";
                        sQry += "           AND U_UseYN = 'Y'";
                        sQry += "           AND U_Char1 = '" + RspCode + "'";
                        sQry += "           AND U_Char2 = '" + TeamCode + "'";
                        sQry += " ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "CntcCode":

                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oForm.Items.Item("CntcName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        
        //public void PS_PP045_MTX01()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_PP045_MTX01()
        //	//해당모듈 : PS_PP045
        //	//기능 : 데이터 조회
        //	//인수 : 없음
        //	//반환값 : 없음
        //	//특이사항 : 없음
        //	//******************************************************************************
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;
        //	string sQry = null;
        //	short ErrNum = 0;

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string BPLId = null;
        //	//사업장
        //	string TeamCode = null;
        //	//팀
        //	string RspCode = null;
        //	//담당
        //	string ClsCode = null;
        //	//반
        //	string OrdGbn = null;
        //	string DocDateFr = null;
        //	string DocDateTo = null;
        //	string CntcCode = null;
        //	//사번

        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	RspCode = Strings.Trim(oForm.Items.Item("RspCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ClsCode = Strings.Trim(oForm.Items.Item("ClsCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	OrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocDateFr = Strings.Trim(oForm.Items.Item("DocDateFr").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocDateTo = Strings.Trim(oForm.Items.Item("DocDateTo").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	CntcCode = Strings.Trim(oForm.Items.Item("CntcCode").Specific.Value);

        //	if (string.IsNullOrEmpty(BPLId))
        //		BPLId = "%";
        //	if (string.IsNullOrEmpty(OrdGbn))
        //		OrdGbn = "%";
        //	if (string.IsNullOrEmpty(DocDateFr))
        //		DocDateFr = "18990101";
        //	if (string.IsNullOrEmpty(DocDateTo))
        //		DocDateTo = "20991231";
        //	if (string.IsNullOrEmpty(CntcCode))
        //		CntcCode = "%";

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //	oForm.Freeze(true);

        //	sQry = "         EXEC [PS_PP045_01] '";
        //	sQry = sQry + BPLId + "','";
        //	sQry = sQry + TeamCode + "','";
        //	sQry = sQry + RspCode + "','";
        //	sQry = sQry + ClsCode + "','";
        //	sQry = sQry + OrdGbn + "','";
        //	sQry = sQry + DocDateFr + "','";
        //	sQry = sQry + DocDateTo + "','";
        //	sQry = sQry + CntcCode + "'";
        //	oRecordSet01.DoQuery(sQry);

        //	oMat01.Clear();
        //	oDS_PS_USERDS01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	if ((oRecordSet01.RecordCount == 0))
        //	{

        //		ErrNum = 1;

        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //		//        Call PS_PP045_Add_MatrixRow(0, True)
        //		goto PS_PP045_MTX01_Error;

        //		return;
        //	}

        //	for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //	{
        //		if (i + 1 > oDS_PS_USERDS01.Size)
        //		{
        //			oDS_PS_USERDS01.InsertRecord((i));
        //		}

        //		oMat01.AddRow();
        //		oDS_PS_USERDS01.Offset = i;

        //		oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //		oDS_PS_USERDS01.SetValue("U_ColDt01", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("DocDate").Value), "YYYYMMDD"));
        //		oDS_PS_USERDS01.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("BPLId").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("OrdType").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("OrdGbn").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("CntcCode").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColReg06", i, Strings.Trim(oRecordSet01.Fields.Item("FullName").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColQty01", i, Strings.Trim(oRecordSet01.Fields.Item("YTime").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColQty02", i, Strings.Trim(oRecordSet01.Fields.Item("WorkTime").Value));
        //		oDS_PS_USERDS01.SetValue("U_ColQty03", i, Strings.Trim(oRecordSet01.Fields.Item("Diff").Value));

        //		oRecordSet01.MoveNext();
        //		ProgBar01.Value = ProgBar01.Value + 1;
        //		ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

        //	}

        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	ProgBar01.Stop();
        //	oForm.Freeze(false);

        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	return;
        //	PS_PP045_MTX01_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//    ProgBar01.Stop
        //	oForm.Freeze(false);
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	if (ErrNum == 1)
        //	{
        //		MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
        //	}
        //	else
        //	{
        //		MDC_Com.MDC_GF_Message(ref "PS_PP045_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}


        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	int ErrNum = 0;
        //	object TempForm01 = null;
        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //	string ItemType = null;
        //	string RequestDate = null;
        //	string Size = null;
        //	string ItemCode = null;
        //	string ItemName = null;
        //	string Unit = null;
        //	string DueDate = null;
        //	string RequestNo = null;
        //	int Qty = 0;
        //	decimal Weight = default(decimal);
        //	string RFC_Sender = null;
        //	double Calculate_Weight = 0;
        //	int Seq = 0;

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true))
        //	{
        //		switch (pval.EventType)
        //		{
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1

        //				if (pval.ItemUID == "Btn01")
        //				{

        //					this.PS_PP045_MTX01();

        //				}
        //				break;

        //			//et_KEY_DOWN ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				if (pval.CharPressed == 9)
        //				{
        //					if (pval.ItemUID == "CntcCode")
        //					{
        //						//UPGRADE_WARNING: oForm.Items(CntcCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
        //						{
        //							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //							BubbleEvent = false;
        //						}
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				break;
        //		}
        //		////BeforeAction = False
        //	}
        //	else if ((pval.BeforeAction == false))
        //	{
        //		switch (pval.EventType)
        //		{
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				break;
        //			//et_COMBO_SELECT ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				if (pval.ItemChanged == true)
        //				{
        //					//If pval.ItemUID = "BPLId" Or pval.ItemUID = "OrdGbn" Then
        //					PS_PP045_FlushToItemValue(pval.ItemUID);
        //					//End If
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			//et_MATRIX_LINK_PRESSED /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				if (pval.ItemUID == "Mat01" & pval.ColUID == "DocEntry")
        //				{
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (Strings.Trim(oMat01.Columns.Item("BPLId").Cells.Item(pval.Row).Specific.Value) == "3")
        //					{
        //						TempForm01 = new PS_PP043();
        //					}
        //					else
        //					{
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Strings.Trim(oMat01.Columns.Item("OrdGbn").Cells.Item(pval.Row).Specific.Value) == "104" | Strings.Trim(oMat01.Columns.Item("OrdGbn").Cells.Item(pval.Row).Specific.Value) == "107")
        //						{
        //							TempForm01 = new PS_PP041();
        //						}
        //						else
        //						{
        //							TempForm01 = new PS_PP040();
        //						}
        //					}
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					TempForm01.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pval.Row).Specific.Value);

        //					//UPGRADE_NOTE: TempForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					TempForm01 = null;
        //				}
        //				break;
        //			//et_VALIDATE ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				if (pval.ItemChanged == true)
        //				{
        //					//If pval.ItemUID = "DocDateFr" Or pval.ItemUID = "DocDateTo" Or pval.ItemUID = "CntcCode" Then
        //					PS_PP045_FlushToItemValue(pval.ItemUID);
        //					//End If
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				//UPGRADE_NOTE: oDS_PS_USERDS01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oDS_PS_USERDS01 = null;
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	if (ErrNum == 101)
        //	{
        //		ErrNum = 0;
        //		MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		BubbleEvent = false;
        //	}
        //	else
        //	{
        //		MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true))
        //	{
        //		switch (pval.MenuUID)
        //		{
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //		////BeforeAction = False
        //	}
        //	else if ((pval.BeforeAction == false))
        //	{
        //		switch (pval.MenuUID)
        //		{
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true))
        //	{
        //		switch (BusinessObjectInfo.EventType)
        //		{
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //		////BeforeAction = False
        //	}
        //	else if ((BusinessObjectInfo.BeforeAction == false))
        //	{
        //		switch (BusinessObjectInfo.EventType)
        //		{
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if ((eventInfo.BeforeAction == true))
        //	{
        //		////작업
        //	}
        //	else if ((eventInfo.BeforeAction == false))
        //	{
        //		////작업
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion




    }
}
