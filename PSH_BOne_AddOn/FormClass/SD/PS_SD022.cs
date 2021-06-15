using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 판매계획 대비 실적 현황
    /// </summary>
    internal class PS_SD022 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_SD022L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD022.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD022_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD022");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry"

                oForm.Freeze(true);
                PS_SD022_CreateItems();
                PS_SD022_SetComboBox();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1285", false); //복원
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", false); //행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
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
        private void PS_SD022_CreateItems()
        {
            try
            {
                oDS_PS_SD022L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
                oForm.DataSources.UserDataSources.Item("StdYear").Value = DateTime.Now.ToString("yyyy");

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

                //팀
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                //사번
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //성명
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
        private void PS_SD022_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        
            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        

        #region PS_SD022_AddMatrixRow
        //		public void PS_SD022_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////행추가여부
        //			if (RowIserted == false) {
        //				oDS_PS_SD022L.InsertRecord((oRow));
        //			}

        //			oMat01.AddRow();
        //			oDS_PS_SD022L.Offset = oRow;
        //			oDS_PS_SD022L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

        //			oMat01.LoadFromDataSource();
        //			return;
        //			PS_SD022_AddMatrixRow_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Com.MDC_GF_Message(ref "PS_SD022_AddMatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion

        #region PS_SD022_MTX01
        //		public void PS_SD022_MTX01()
        //		{
        //			//******************************************************************************
        //			//Function ID : PS_SD022_MTX01()
        //			//해당모듈 : PS_SD022
        //			//기능 : 데이터 조회
        //			//인수 : 없음
        //			//반환값 : 없음
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			string sQry = null;
        //			short ErrNum = 0;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string StdYear = null;
        //			//기준년도
        //			string BPLID = null;
        //			//사업장
        //			string TeamCode = null;
        //			//소속팀
        //			string RspCode = null;
        //			//소속담당
        //			string CntcCode = null;
        //			//영업담당자사번

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			StdYear = Strings.Trim(oForm.Items.Item("StdYear").Specific.Value);
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BPLID = (Strings.Trim(oForm.Items.Item("BPLID").Specific.Value) == "%" ? "" : Strings.Trim(oForm.Items.Item("BPLID").Specific.Value));
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.Value);
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			RspCode = Strings.Trim(oForm.Items.Item("RspCode").Specific.Value);
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CntcCode = Strings.Trim(oForm.Items.Item("CntcCode").Specific.Value);

        //			SAPbouiCOM.ProgressBar ProgBar01 = null;
        //			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //			oForm.Freeze(true);

        //			sQry = "EXEC [PS_SD022_01] '";
        //			sQry = sQry + StdYear + "','";
        //			sQry = sQry + BPLID + "','";
        //			sQry = sQry + TeamCode + "','";
        //			sQry = sQry + RspCode + "','";
        //			sQry = sQry + CntcCode + "'";
        //			oRecordSet01.DoQuery(sQry);

        //			oMat01.Clear();
        //			oDS_PS_SD022L.Clear();
        //			oMat01.FlushToDataSource();
        //			oMat01.LoadFromDataSource();

        //			if ((oRecordSet01.RecordCount == 0)) {

        //				ErrNum = 1;

        //				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //				//        Call PS_SD022_AddMatrixRow(0, True)
        //				goto PS_SD022_MTX01_Error;

        //				return;
        //			}

        //			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
        //				if (i + 1 > oDS_PS_SD022L.Size) {
        //					oDS_PS_SD022L.InsertRecord((i));
        //				}

        //				oMat01.AddRow();
        //				oDS_PS_SD022L.Offset = i;

        //				oDS_PS_SD022L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //				oDS_PS_SD022L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("ClassNm").Value));
        //				//구분
        //				oDS_PS_SD022L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("Class2").Value));
        //				//판매
        //				oDS_PS_SD022L.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("Month01").Value));
        //				//1월
        //				oDS_PS_SD022L.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("Month02").Value));
        //				//2월
        //				oDS_PS_SD022L.SetValue("U_ColSum03", i, Strings.Trim(oRecordSet01.Fields.Item("Month03").Value));
        //				//3월
        //				oDS_PS_SD022L.SetValue("U_ColSum04", i, Strings.Trim(oRecordSet01.Fields.Item("Month04").Value));
        //				//4월
        //				oDS_PS_SD022L.SetValue("U_ColSum05", i, Strings.Trim(oRecordSet01.Fields.Item("Month05").Value));
        //				//5월
        //				oDS_PS_SD022L.SetValue("U_ColSum06", i, Strings.Trim(oRecordSet01.Fields.Item("Month06").Value));
        //				//6월
        //				oDS_PS_SD022L.SetValue("U_ColSum07", i, Strings.Trim(oRecordSet01.Fields.Item("Month07").Value));
        //				//7월
        //				oDS_PS_SD022L.SetValue("U_ColSum08", i, Strings.Trim(oRecordSet01.Fields.Item("Month08").Value));
        //				//8월
        //				oDS_PS_SD022L.SetValue("U_ColSum09", i, Strings.Trim(oRecordSet01.Fields.Item("Month09").Value));
        //				//9월
        //				oDS_PS_SD022L.SetValue("U_ColSum10", i, Strings.Trim(oRecordSet01.Fields.Item("Month10").Value));
        //				//10월
        //				oDS_PS_SD022L.SetValue("U_ColSum11", i, Strings.Trim(oRecordSet01.Fields.Item("Month11").Value));
        //				//11월
        //				oDS_PS_SD022L.SetValue("U_ColSum12", i, Strings.Trim(oRecordSet01.Fields.Item("Month12").Value));
        //				//12월

        //				oRecordSet01.MoveNext();
        //				ProgBar01.Value = ProgBar01.Value + 1;
        //				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

        //			}

        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			ProgBar01.Stop();
        //			oForm.Freeze(false);

        //			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgBar01 = null;
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			return;
        //			PS_SD022_MTX01_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//    ProgBar01.Stop
        //			oForm.Freeze(false);
        //			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgBar01 = null;
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;

        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_SD022_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region PS_SD022_FlushToItemValue
        //		private void PS_SD022_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			short ErrNum = 0;
        //			string sQry = null;
        //			string ItemCode = null;
        //			short Qty = 0;
        //			double Calculate_Weight = 0;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string BPLID = null;
        //			string TeamCode = null;

        //			switch (oUID) {

        //				case "BPLID":

        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					BPLID = Strings.Trim(oForm.Items.Item("BPLID").Specific.Value);

        //					//UPGRADE_WARNING: oForm.Items(TeamCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0) {
        //						//UPGRADE_WARNING: oForm.Items(TeamCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1) {
        //							//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
        //						}
        //					}

        //					//부서콤보세팅
        //					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
        //					sQry = "            SELECT      U_Code AS [Code],";
        //					sQry = sQry + "                 U_CodeNm As [Name]";
        //					sQry = sQry + "  FROM       [@PS_HR200L]";
        //					sQry = sQry + "  WHERE      Code = '1'";
        //					sQry = sQry + "                 AND U_UseYN = 'Y'";
        //					sQry = sQry + "                 AND U_Char2 = '" + BPLID + "'";
        //					sQry = sQry + "  ORDER BY  U_Seq";
        //					MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("TeamCode").Specific), ref sQry, ref "", ref false, ref false);
        //					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //					break;

        //				case "TeamCode":

        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.Value);

        //					//UPGRADE_WARNING: oForm.Items(RspCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0) {
        //						//UPGRADE_WARNING: oForm.Items(RspCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1) {
        //							//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
        //						}
        //					}

        //					//담당콤보세팅
        //					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
        //					sQry = "            SELECT      U_Code AS [Code],";
        //					sQry = sQry + "                 U_CodeNm As [Name]";
        //					sQry = sQry + "  FROM       [@PS_HR200L]";
        //					sQry = sQry + "  WHERE      Code = '2'";
        //					sQry = sQry + "                 AND U_UseYN = 'Y'";
        //					sQry = sQry + "                 AND U_Char1 = '" + TeamCode + "'";
        //					sQry = sQry + "  ORDER BY  U_Seq";
        //					MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("RspCode").Specific), ref sQry, ref "", ref false, ref false);
        //					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //					break;

        //				case "CntcCode":

        //					//UPGRADE_WARNING: oForm.Items(CntcName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("CntcName").Specific.Value = MDC_GetData.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value + "'");
        //					//성명
        //					break;

        //			}

        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			return;
        //			PS_SD022_FlushToItemValue_Error:

        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;

        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "시각은 숫자만 입력이 가능합니다.", ref "E");
        //			} else if (ErrNum == 2) {
        //				MDC_Com.MDC_GF_Message(ref "시각(시)는 24미만의 값만 입력이 가능합니다.", ref "E");
        //			} else if (ErrNum == 3) {
        //				MDC_Com.MDC_GF_Message(ref "시각(분)은 60미만의 값만 입력이 가능합니다.", ref "E");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_SD022_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region Raise_ItemEvent
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //					////8
        //					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //					////10
        //					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //					////18
        //					break;
        //				////et_FORM_ACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //					////19
        //					break;
        //				////et_FORM_DEACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //					////20
        //					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //					////27
        //					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;
        //				////et_LOST_FOCUS
        //				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //					////17
        //					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //			}
        //			return;
        //			Raise_ItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_MenuEvent
        //		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						///추가버튼 클릭시 메트릭스 insertrow

        //						//                oMat01.Clear
        //						//                oMat01.FlushToDataSource
        //						//                oMat01.LoadFromDataSource

        //						//                oForm.Mode = fm_ADD_MODE
        //						//                BubbleEvent = False

        //						//oForm.Items("GCode").Click ct_Regular


        //						return;

        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					////Call PS_SD022_FormItemEnabled '//UDO방식
        //					case "1282":
        //						//추가
        //						break;
        //					//                oMat01.Clear
        //					//                oDS_PS_SD022H.Clear

        //					//                Call PS_SD022_LoadCaption
        //					//                Call PS_SD022_FormItemEnabled
        //					////Call PS_SD022_FormItemEnabled '//UDO방식
        //					////Call PS_SD022_AddMatrixRow(0, True) '//UDO방식
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //					////Call PS_SD022_FormItemEnabled
        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}
        //			if (pval.ItemUID == "Mat01") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //				if (pval.ItemUID == "PS_SD022") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}

        //				///조회
        //				if (pval.ItemUID == "BtnSearch") {

        //					PS_SD022_MTX01();

        //				//출력
        //				} else if (pval.ItemUID == "BtnPrint") {

        //					PS_SD022_Print_Report01();

        //				}

        //			} else if (pval.BeforeAction == false) {
        //				if (pval.ItemUID == "PS_SD022") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "CntcCode", "");
        //				////사용자값활성
        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_CLICK
        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //				if (pval.ItemUID == "Mat01") {

        //					if (pval.Row > 0) {

        //						oMat01.SelectRow(pval.Row, true, false);

        //					}

        //				}

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_VALIDATE
        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //				if (pval.ItemChanged == true) {

        //					if ((pval.ItemUID == "Mat01")) {
        //					} else {

        //						PS_SD022_FlushToItemValue(pval.ItemUID);
        //				}

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_RESIZE
        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				PS_SD022_ResizeForm();
        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.ItemUID == "Mat01") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}
        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //				SubMain.RemoveForms(oFormUniqueID);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			if ((oLastColRow01 > 0)) {
        //				if (pval.BeforeAction == true) {
        //					//            If (PS_SD022_Validate("행삭제") = False) Then
        //					//                BubbleEvent = False
        //					//                Exit Sub
        //					//            End If
        //					////행삭제전 행삭제가능여부검사
        //				} else if (pval.BeforeAction == false) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					oDS_PS_SD022H.RemoveRecord(oDS_PS_SD022H.Size - 1);
        //					oMat01.LoadFromDataSource();
        //					if (oMat01.RowCount == 0) {
        //						PS_SD022_AddMatrixRow(0);
        //					} else {
        //						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD022H.GetValue("U_CntcCode", oMat01.RowCount - 1)))) {
        //							PS_SD022_AddMatrixRow(oMat01.RowCount);
        //						}
        //					}
        //				}
        //			}
        //			return;
        //			Raise_EVENT_ROW_DELETE_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion




        #region PS_SD022_ResizeForm
        //		private void PS_SD022_ResizeForm()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oMat01.AutoResizeColumns();

        //			return;
        //			PS_SD022_ResizeForm_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SD022_ResizeForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region PS_SD022_Print_Report01
        //		private void PS_SD022_Print_Report01()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string DocNum = null;
        //			string WinTitle = null;
        //			string ReportName = null;
        //			string sQry = null;

        //			short i = 0;
        //			short ErrNum = 0;
        //			string Sub_sQry = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			MDC_PS_Common.ConnectODBC();

        //			string CntcCode = null;
        //			string BPLID = null;
        //			//사업장
        //			string ZGODAYF = null;
        //			//기준일자F
        //			string ZGODAYT = null;
        //			//기준일자T

        //			CntcCode = MDC_PS_Common.User_MSTCOD();
        //			//조회자사번
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BPLID = Strings.Trim(oForm.Items.Item("SBPLID").Specific.Value);
        //			//사업장
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ZGODAYF = Strings.Trim(oForm.Items.Item("SZGODAYF").Specific.Value);
        //			//기준일자F
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ZGODAYT = Strings.Trim(oForm.Items.Item("SZGODAYT").Specific.Value);
        //			//기준일자T

        //			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			WinTitle = "[PS_SD022] 레포트";
        //			ReportName = "PS_SD022_01.rpt";
        //			MDC_Globals.gRpt_Formula = new string[3];
        //			MDC_Globals.gRpt_Formula_Value = new string[3];
        //			MDC_Globals.gRpt_SRptSqry = new string[2];
        //			MDC_Globals.gRpt_SRptName = new string[2];
        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //			//// Formula 수식필드

        //			//// SubReport


        //			MDC_Globals.gRpt_SFormula[1, 1] = "";
        //			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //			/// Procedure 실행"
        //			sQry = "         EXEC [PS_SD022_01] '";
        //			sQry = sQry + CntcCode + "','";
        //			sQry = sQry + BPLID + "','";
        //			sQry = sQry + ZGODAYF + "','";
        //			sQry = sQry + ZGODAYT + "'";

        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				ErrNum = 1;
        //				goto Print_Query_Error;
        //			}

        //			/// Action (sub_query가 있을때는 'Y'로...)/
        //			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			Print_Query_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_SD022_Print_Report01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion
    }
}
