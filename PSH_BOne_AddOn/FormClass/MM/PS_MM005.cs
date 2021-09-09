using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 구매요청
	/// </summary>
	internal class PS_MM005 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_MM005H; //등록헤더
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_RightClick_CgNum;
		
        //모드 변경에 따른 조회조건 저장용 클래스
        private class SearchData
        {
            private string ordType;
            private string bPLId;
            private string cntcCode;
            private string cntcName;
            private string deptCode;
            private string docDateFr;
            private string docDateTo;
            private string cgNumFr;
            private string cgNumTo;
            private string itemCode;
            private string itmBSort;
            private string itmMSort;
            private string itemType;
            private string oKYN;

            //Setter
            public void SetOrdType(string ordType) { this.ordType = ordType; }
            public void SetBPLID(string bPLId) { this.bPLId = bPLId; }
            public void SetCntcCode(string cntcCode) { this.cntcCode = cntcCode; }
            public void SetCntcName(string cntcName) { this.cntcName = cntcName; }
            public void SetDeptCode(string deptCode) { this.deptCode = deptCode; }
            public void SetDocDateFr(string docDateFr) { this.docDateFr = docDateFr; }
            public void SetDocDateTo(string docDateTo) { this.docDateTo = docDateTo; }
            public void SetCgNumFr(string cgNumFr) { this.cgNumFr = cgNumFr; }
            public void SetCgNumTo(string cgNumTo) { this.cgNumTo = cgNumTo; }
            public void SetItemCode(string itemCode) { this.itemCode = itemCode; }
            public void SetItmBSort(string itmBSort) { this.itmBSort = itmBSort; }
            public void SetItmMSort(string itmMSort) { this.itmMSort = itmMSort; }            
            public void SetItemType(string itemType) { this.itemType = itemType; }
            public void SetOKYN(string oKYN) { this.oKYN = oKYN; }

            //Getter
            public string GetOrdType() { return ordType; }
            public string GetBPLID() { return bPLId; }
            public string GetCntcCode() { return cntcCode; }
            public string GetCntcName() { return cntcName; }
            public string GetDeptCode() { return deptCode; }
            public string GetDocDateFr() { return docDateFr; }
            public string GetDocDateTo() { return docDateTo; }
            public string GetCgNumFr() { return cgNumFr; }
            public string GetCgNumTo() { return cgNumTo; }
            public string GetItemCode() { return itemCode; }
            public string GetItmBSort() { return itmBSort; }
            public string GetItmMSort() { return itmMSort; }
            public string GetItemType() { return itemType; }
            public string GetOKYN() { return oKYN; }
        }

        SearchData searchData = new SearchData();

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM005.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM005_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM005");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				
				oForm.Freeze(true);
				PS_MM005_CreateItems();
				PS_MM005_SetComboBox();
				PS_MM005_Initialize();
				PS_MM005_AddMatrixRow(0, true);
				PS_MM005_LoadCaption();
				PS_MM005_EnableFormItem();

				oForm.EnableMenu("1281", true); //찾기
				oForm.EnableMenu("1282", true); //추가
				oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1299", true); //행닫기
				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", false); //취소
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
        private void PS_MM005_CreateItems()
        {
            try
            {
                oDS_PS_MM005H = oForm.DataSources.DBDataSources.Item("@PS_MM005H");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //요청일자(FR)
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                
                //요청일자(TO)
                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                
                //대분류
                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                //청구인(사번)
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //청구인(성명)
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //청구번호(FR)
                oForm.DataSources.UserDataSources.Add("CgNumFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CgNumFr").Specific.DataBind.SetBound(true, "", "CgNumFr");

                //청구번호(TO)
                oForm.DataSources.UserDataSources.Add("CgNumTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CgNumTo").Specific.DataBind.SetBound(true, "", "CgNumTo");

                //중분류
                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

                //품목구분
                oForm.DataSources.UserDataSources.Add("OrdType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("OrdType").Specific.DataBind.SetBound(true, "", "OrdType");

                //품목코드
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목타입
                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                //청구부서
                oForm.DataSources.UserDataSources.Add("DeptCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("DeptCode").Specific.DataBind.SetBound(true, "", "DeptCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //결재여부
                oForm.DataSources.UserDataSources.Add("OKYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("OKYN").Specific.DataBind.SetBound(true, "", "OKYN");

                //본수합계
                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                //수량/중량합계
                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                //제품BOM사용
                oForm.DataSources.UserDataSources.Add("BOM_CHECK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("BOM_CHECK").Specific.DataBind.SetBound(true, "", "BOM_CHECK");

                oMat01.Columns.Item("PP030DL").Editable = false;
                oMat01.Columns.Item("ItemCode").Editable = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 설정
        /// </summary>
        private void PS_MM005_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //품목구분
                sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("OrdType").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("OrdType").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("OrdType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //청구부서
                sQry = "SELECT Code, Name From [OUDP] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("DeptCode").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("DeptCode").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //사용처
                sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' AND Active = 'Y' Order by PrcCode";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("UseDept").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //대분류
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
                oRecordSet01.DoQuery(sQry);
                oForm.Items.Item("ItmBSort").Specific.ValidValues.Add("%", "전체");
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //품목타입
                sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code";
                oRecordSet01.DoQuery(sQry);
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("ItemType").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //결재여부
                oForm.Items.Item("OKYN").Specific.ValidValues.Add("Y", "결재");
                oForm.Items.Item("OKYN").Specific.ValidValues.Add("N", "미결재");
                oForm.Items.Item("OKYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oMat01.Columns.Item("DivideYN").ValidValues.Add("N", "분할입고");
                oMat01.Columns.Item("DivideYN").ValidValues.Add("Y", "일반입고");

                //외주사유코드
                sQry = "  SELECT    T1.U_Minor AS [Code],";
                sQry += "           T1.U_CdName AS [Value]";
                sQry += " FROM      [@PS_SY001H] AS T0";
                sQry += "           INNER JOIN";
                sQry += "           [@PS_SY001L] AS T1";
                sQry += "               ON T0.Code = T1.Code";
                sQry += " WHERE     T0.Code = 'P201'";
                sQry += "           AND T1.U_UseYN = 'Y'";
                sQry += " ORDER BY  T1.U_Seq";

                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("OutCode").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //청구사유(라인)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'P203'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += "           AND U_Minor <> 'A'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RCode"), sQry, "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PS_MM005_Initialize()
        {
            string appUserYN; //등록권한 보유 여부
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.DataSources.UserDataSources.Item("BPLId").Value = dataHelpClass.User_BPLID(); //사업장
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM01"); //요청일자(Fr)
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd"); //요청일자(To)
                oForm.DataSources.UserDataSources.Item("CntcCode").Value = dataHelpClass.User_MSTCOD(); //사번
                oForm.DataSources.UserDataSources.Item("CntcName").Value = dataHelpClass.User_MSTNAM(); //성명
                oForm.DataSources.UserDataSources.Item("DeptCode").Value = dataHelpClass.User_DeptCode(); //부서

                //보유한 권한별 Btn01 설정
                appUserYN = dataHelpClass.Get_ReData("U_UserID", "U_UserID", "[@PS_SY005L]", "'" + PSH_Globals.oCompany.UserName + "'", " AND Code = 'MM005'"); 

                if (!string.IsNullOrEmpty(appUserYN))
                {
                    oForm.Items.Item("Btn01").Enabled = true;
                }
                else
                {
                    oForm.Items.Item("Btn01").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_MM005_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_MM005H.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM005H.Offset = oRow;
                oDS_PS_MM005H.SetValue("DocNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 모드 상태에 따른 버튼 Caption 설정
        /// </summary>
        private void PS_MM005_LoadCaption()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "추가";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "확인";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "갱신";
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에 따른 아이템설정
        /// </summary>
        private void PS_MM005_EnableFormItem()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DeptCode").Enabled = true;
                    oForm.Items.Item("DocDateFr").Enabled = false;
                    oForm.Items.Item("DocDateTo").Enabled = false;
                    oForm.Items.Item("CgNumFr").Enabled = false;
                    oForm.Items.Item("CgNumTo").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("ItmBSort").Enabled = false;
                    oForm.Items.Item("ItmMSort").Enabled = false;
                    oForm.Items.Item("OKYN").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;

                    oMat01.Columns.Item("PP030DL").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemName").Editable = false;
                    oMat01.Columns.Item("Qty").Editable = true;
                    oMat01.Columns.Item("Weight").Editable = true;
                    oMat01.Columns.Item("BPLId").Editable = false;
                    oMat01.Columns.Item("CgNum").Editable = false;
                    oMat01.Columns.Item("DocDate").Editable = true;
                    oMat01.Columns.Item("DueDate").Editable = true;
                    oMat01.Columns.Item("CntcCode").Editable = false;
                    oMat01.Columns.Item("CntcName").Editable = false;
                    oMat01.Columns.Item("DeptCode").Editable = false;
                    oMat01.Columns.Item("UseDept").Editable = true;
                    oMat01.Columns.Item("Auto").Editable = true;
                    oMat01.Columns.Item("QCYN").Editable = true;
                    oMat01.Columns.Item("Note").Editable = true;
                    oMat01.Columns.Item("Comments").Editable = true;
                    oMat01.Columns.Item("IvQty").Editable = false;
                    oMat01.Columns.Item("IvWeight").Editable = false;
                    oMat01.Columns.Item("OrdType").Editable = false;
                    oMat01.Columns.Item("OKYN").Editable = false;
                    oMat01.Columns.Item("OKDate").Editable = false;
                    oMat01.Columns.Item("UpdtUser").Visible = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DeptCode").Enabled = true;
                    oForm.Items.Item("DocDateFr").Enabled = true;
                    oForm.Items.Item("DocDateTo").Enabled = true;
                    oForm.Items.Item("CgNumFr").Enabled = true;
                    oForm.Items.Item("CgNumTo").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("ItmMSort").Enabled = true;
                    oForm.Items.Item("OKYN").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;

                    oMat01.Clear();

                    oMat01.Columns.Item("PP030DL").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemName").Editable = false;
                    oMat01.Columns.Item("Qty").Editable = true;
                    oMat01.Columns.Item("Weight").Editable = true;
                    oMat01.Columns.Item("BPLId").Editable = false;
                    oMat01.Columns.Item("CgNum").Editable = false;
                    oMat01.Columns.Item("DocDate").Editable = true;
                    oMat01.Columns.Item("DueDate").Editable = true;
                    oMat01.Columns.Item("CntcCode").Editable = false;
                    oMat01.Columns.Item("CntcName").Editable = false;
                    oMat01.Columns.Item("DeptCode").Editable = false;
                    oMat01.Columns.Item("UseDept").Editable = true;
                    oMat01.Columns.Item("Auto").Editable = true;
                    oMat01.Columns.Item("QCYN").Editable = true;
                    oMat01.Columns.Item("Note").Editable = true;
                    oMat01.Columns.Item("Comments").Editable = true;
                    oMat01.Columns.Item("IvQty").Editable = false;
                    oMat01.Columns.Item("IvWeight").Editable = false;
                    oMat01.Columns.Item("OrdType").Editable = false;
                    oMat01.Columns.Item("OKYN").Editable = false;
                    oMat01.Columns.Item("OKDate").Editable = false;
                    oMat01.Columns.Item("UpdtUser").Visible = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DeptCode").Enabled = true;
                    oForm.Items.Item("DocDateFr").Enabled = false;
                    oForm.Items.Item("DocDateTo").Enabled = false;
                    oForm.Items.Item("CgNumFr").Enabled = false;
                    oForm.Items.Item("CgNumTo").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("ItmBSort").Enabled = false;
                    oForm.Items.Item("ItmMSort").Enabled = false;
                    oForm.Items.Item("OKYN").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;

                    if (oForm.DataSources.UserDataSources.Item("OKYN").Value == "Y")
                    {
                        oMat01.Columns.Item("PP030DL").Editable = false;
                        oMat01.Columns.Item("ItemCode").Editable = false;
                        oMat01.Columns.Item("ItemName").Editable = false;
                        oMat01.Columns.Item("Qty").Editable = false;
                        oMat01.Columns.Item("Weight").Editable = false;
                        oMat01.Columns.Item("BPLId").Editable = false;
                        oMat01.Columns.Item("CgNum").Editable = false;
                        oMat01.Columns.Item("DocDate").Editable = false;
                        oMat01.Columns.Item("DueDate").Editable = false;
                        oMat01.Columns.Item("CntcCode").Editable = false;
                        oMat01.Columns.Item("CntcName").Editable = false;
                        oMat01.Columns.Item("DeptCode").Editable = false;
                        oMat01.Columns.Item("UseDept").Editable = false;
                        oMat01.Columns.Item("Auto").Editable = false;
                        oMat01.Columns.Item("QCYN").Editable = false;
                        oMat01.Columns.Item("Note").Editable = false;
                        oMat01.Columns.Item("Comments").Editable = false;
                        oMat01.Columns.Item("IvQty").Editable = false;
                        oMat01.Columns.Item("IvWeight").Editable = false;
                        oMat01.Columns.Item("OrdType").Editable = false;
                        oMat01.Columns.Item("OKYN").Editable = false;
                        oMat01.Columns.Item("OKDate").Editable = false;
                        oMat01.Columns.Item("Status").Editable = false;
                        oMat01.Columns.Item("UpdtUser").Visible = false;
                    }
                    else
                    {
                        oMat01.Columns.Item("PP030DL").Editable = false;
                        oMat01.Columns.Item("ItemCode").Editable = true;
                        oMat01.Columns.Item("ItemName").Editable = false;
                        oMat01.Columns.Item("Qty").Editable = true;
                        oMat01.Columns.Item("Weight").Editable = true;
                        oMat01.Columns.Item("BPLId").Editable = false;
                        oMat01.Columns.Item("CgNum").Editable = false;
                        oMat01.Columns.Item("DocDate").Editable = true;
                        oMat01.Columns.Item("DueDate").Editable = true;
                        oMat01.Columns.Item("CntcCode").Editable = false;
                        oMat01.Columns.Item("CntcName").Editable = false;
                        oMat01.Columns.Item("DeptCode").Editable = false;
                        oMat01.Columns.Item("UseDept").Editable = true;
                        oMat01.Columns.Item("Auto").Editable = true;
                        oMat01.Columns.Item("QCYN").Editable = true;

                        oMat01.Columns.Item("Note").Editable = true;
                        oMat01.Columns.Item("Comments").Editable = true;
                        oMat01.Columns.Item("IvQty").Editable = false;
                        oMat01.Columns.Item("IvWeight").Editable = false;
                        oMat01.Columns.Item("OrdType").Editable = false;
                        oMat01.Columns.Item("OKYN").Editable = false;
                        oMat01.Columns.Item("OKDate").Editable = false;
                        oMat01.Columns.Item("Status").Editable = true;
                        oMat01.Columns.Item("UpdtUser").Visible = false;
                    }
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
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM005_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            double SumQty = 0;
            double SumWeight = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("CntcName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "ItemCode":
                        sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "Mat01":
                        if (oCol == "ItemCode")
                        { 
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM005_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "60")
                            {
                                sQry = "Select U_ItemName, U_OutUnit From [@PS_MM005H] Where U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                                oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                oMat01.Columns.Item("OutUnit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                            }
                            else
                            {
                                sQry = "Select ItemName, FrgnName, BuyUnitMsr From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                                oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                oMat01.Columns.Item("OutUnit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(2).Value.ToString().Trim();
                            
                                sQry = "Select Sum(OnHand) From OITW Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                                oMat01.Columns.Item("IvQty").Cells.Item(oRow).Specific.Value = dataHelpClass.Calculate_Qty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim(), Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()));
                                oMat01.Columns.Item("IvWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }

                            oMat01.Columns.Item("DocDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("Auto").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("QCYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("OKYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("DivideYN").Cells.Item(oRow).Specific.Select("N");

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "ItemName")
                        {
                            if (oRow == oMat01.RowCount)
                            {
                                oMat01.FlushToDataSource();
                                PS_MM005_AddMatrixRow(oMat01.RowCount, false);
                            }
                            oMat01.Columns.Item("Auto").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("QCYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("OKYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("DivideYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "PP030DL")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("PP030DL").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM005_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            sQry = "EXEC [PS_MM005_02] '";
                            sQry += oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "', '";
                            sQry += oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() + "', '";
                            sQry += oMat01.Columns.Item("PP030DL").Cells.Item(oRow).Specific.Value.ToString().Trim() + "', '";
                            sQry += "2'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                            oMat01.Columns.Item("OutSize").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("Size").Value.ToString().Trim();
                            oMat01.Columns.Item("DocDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("Auto").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("QCYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("OKYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("DivideYN").Cells.Item(oRow).Specific.Select("N");
                            oMat01.Columns.Item("ProcCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CpCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ProcName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CpName").Value.ToString().Trim();

                            oMat01.AutoResizeColumns();

                            oMat01.Columns.Item("PP030DL").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "Qty")
                        {   
                            oMat01.FlushToDataSource();
                            double calWeight = dataHelpClass.Calculate_Weight(oDS_PS_MM005H.GetValue("U_ItemCode", oRow - 1).ToString().Trim(), Convert.ToInt32(oDS_PS_MM005H.GetValue("U_Qty", oRow - 1)), oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());
                            oDS_PS_MM005H.SetValue("U_Weight", oRow - 1, Convert.ToString(calWeight)); //가독성을 위한 변수(calWeight) 사용
                            oMat01.LoadFromDataSource();

                            for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value.ToString().Trim()))
                                {
                                    SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                                }
                                SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                            }

                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                            oMat01.Columns.Item("Qty").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "Weight")
                        {
                            for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value.ToString().Trim()))
                                {
                                    SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                                }
                                SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                            }
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                            oMat01.Columns.Item("Weight").Cells.Item(oRow).Click();                            
                        }
                        else if (oCol == "CntcCode")
                        {
                            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oMat01.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("CntcName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (oCol == "ProcCode")
                        {
                            sQry = "select U_CpName From [@ps_pp001L] Where U_CpCode = '" + oMat01.Columns.Item("ProcCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ProcName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 매트릭스의 사용처, 납품일자 일괄 수정
        /// </summary>
        /// <returns></returns>
        private bool PS_MM005_SetMatrixData()
        {
            bool returnValue = false;

            try
            {
                oMat01.FlushToDataSource();

                for (int i = 0; i <= oMat01.VisualRowCount - 3; i++)
                {
                    oDS_PS_MM005H.SetValue("U_UseDept", i + 1, oDS_PS_MM005H.GetValue("U_UseDept", 0).ToString().Trim()); //사용처
                    oDS_PS_MM005H.SetValue("U_DueDate", i + 1, oDS_PS_MM005H.GetValue("U_DueDate", 0).ToString().Trim()); //납품일자
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }            

            return returnValue;
        }

        /// <summary>
        /// Matrix 데이터 로드
        /// </summary>
        private void PS_MM005_LoadData()
        {
            string sQry;
            double SumQty = 0;
            double SumWeight = 0;
            string errMessage = string.Empty;
            string ItemType;
            string ItmBSort;
            string CgNumTo;
            string DeptCode;
            string BPLId;
            string OrdType;
            string CntcCode;
            string CgNumFr;
            string ItemCode;
            string ItmMSort;
            string OKYN;
            int Calculate_Qty;
            string DocDateFr;
            string DocDateTo;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                OrdType = string.IsNullOrEmpty(oForm.Items.Item("OrdType").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
                BPLId = string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                CntcCode = string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                DeptCode = string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
                DocDateFr = string.IsNullOrEmpty(oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim()) ? DateTime.Now.AddMonths(-3).ToString("yyyyMM") + "01" : oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
                DocDateTo = string.IsNullOrEmpty(oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim()) ? DocDateTo = DateTime.Now.ToString("yyyyMMdd") : oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
                CgNumFr = string.IsNullOrEmpty(oForm.Items.Item("CgNumFr").Specific.Value.ToString().Trim()) ? "0000000000" : oForm.Items.Item("CgNumFr").Specific.Value.ToString().Trim();
                CgNumTo = string.IsNullOrEmpty(oForm.Items.Item("CgNumTo").Specific.Value.ToString().Trim()) ? "9999999999" : oForm.Items.Item("CgNumTo").Specific.Value.ToString().Trim();
                ItemCode = string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                ItmBSort = string.IsNullOrEmpty(oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                ItmMSort = string.IsNullOrEmpty(oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();
                ItemType = string.IsNullOrEmpty(oForm.Items.Item("ItemType").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
                OKYN = string.IsNullOrEmpty(oForm.Items.Item("OKYN").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("OKYN").Specific.Value.ToString().Trim();
 
                sQry = "EXEC [PS_MM005_01] '";
                sQry += OrdType + "','";
                sQry += BPLId + "','";
                sQry += CntcCode + "','";
                sQry += DeptCode + "','";
                sQry += DocDateFr + "','";
                sQry += DocDateTo + "','";
                sQry += CgNumFr + "','";
                sQry += CgNumTo + "','";
                sQry += ItemCode + "','";
                sQry += ItmBSort + "','";
                sQry += ItmMSort + "','";
                sQry += ItemType + "','";
                sQry += OKYN + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_MM005H.Clear();

                int temp = oRecordSet01.RecordCount;

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_MM005H.Size)
                    {
                        oDS_PS_MM005H.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_MM005H.Offset = i;
                    oDS_PS_MM005H.SetValue("DocNum", i, Convert.ToString(i + 1));
                    oDS_PS_MM005H.SetValue("U_PP030DL", i, oRecordSet01.Fields.Item("U_PP030DL").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_BPLId", i, oRecordSet01.Fields.Item("U_BPLId").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_Qty", i, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_Weight", i, oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_CgNum", i, oRecordSet01.Fields.Item("U_CgNum").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_DocDate", i, oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_MM005H.SetValue("U_DueDate", i, oRecordSet01.Fields.Item("U_DueDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_MM005H.SetValue("U_CntcCode", i, oRecordSet01.Fields.Item("U_CntcCode").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_CntcName", i, oRecordSet01.Fields.Item("U_CntcName").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_DeptCode", i, oRecordSet01.Fields.Item("U_DeptCode").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_UseDept", i, oRecordSet01.Fields.Item("U_UseDept").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_Auto", i, oRecordSet01.Fields.Item("U_Auto").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_QCYN", i, oRecordSet01.Fields.Item("U_QCYN").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_Note", i, oRecordSet01.Fields.Item("U_Note").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OutCode", i, oRecordSet01.Fields.Item("U_OutCode").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OutNote", i, oRecordSet01.Fields.Item("U_OutNote").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OutSize", i, oRecordSet01.Fields.Item("U_OutSize").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OutUnit", i, oRecordSet01.Fields.Item("U_OutUnit").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OrdNum", i, oRecordSet01.Fields.Item("U_OrdNum").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_UpdtUser", i, oRecordSet01.Fields.Item("U_UpdtUser").Value.ToString().Trim());

                    sQry = "Select Sum(OnHand) From OITW Where ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim() + "'";
                    oRecordSet02.DoQuery(sQry);
                    string itemCode = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                    double tempOnHand = Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim());
                    int sumOnHand = Convert.ToInt32(System.Math.Round(tempOnHand, 0));
                    Calculate_Qty = dataHelpClass.Calculate_Qty(itemCode, sumOnHand);

                    oDS_PS_MM005H.SetValue("U_IvQty", i, Convert.ToString(Calculate_Qty));
                    oDS_PS_MM005H.SetValue("U_IvWeight", i, oRecordSet02.Fields.Item(0).Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_OKYN", i, oRecordSet01.Fields.Item("U_OKYN").Value.ToString().Trim());

                    if (oRecordSet01.Fields.Item("U_OKDate").Value.ToString("yyyyMMdd") == "18991230")
                    {
                        oDS_PS_MM005H.SetValue("U_OKDate", i, "");
                    }
                    else
                    {
                        oDS_PS_MM005H.SetValue("U_OKDate", i, oRecordSet01.Fields.Item("U_OKDate").Value.ToString("yyyyMMdd"));
                    }

                    oDS_PS_MM005H.SetValue("U_OrdType", i, oRecordSet01.Fields.Item("U_OrdType").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_ProcCode", i, oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_ProcName", i, oRecordSet01.Fields.Item("U_ProcName").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_Comments", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                    oDS_PS_MM005H.SetValue("U_ImportYN", i, oRecordSet01.Fields.Item("U_ImportYN").Value.ToString().Trim()); //수입품여부
                    oDS_PS_MM005H.SetValue("U_EmergYN", i, oRecordSet01.Fields.Item("U_EmergYN").Value.ToString().Trim()); //긴급여부
                    oDS_PS_MM005H.SetValue("U_RCode", i, oRecordSet01.Fields.Item("U_RCode").Value.ToString().Trim()); //재청구사유
                    oDS_PS_MM005H.SetValue("U_RName", i, oRecordSet01.Fields.Item("U_RName").Value.ToString().Trim()); //재청구사유내용
                    oDS_PS_MM005H.SetValue("U_Status", i, oRecordSet01.Fields.Item("U_Status").Value.ToString().Trim());

                    SumQty += Convert.ToDouble(oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    SumWeight += Convert.ToDouble(oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중";
                }

                oForm.Items.Item("SumQty").Specific.Value = SumQty;
                oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            }
        }

        /// <summary>
        /// 품목마스터 승인여부 조회
        /// </summary>
        /// <param name="pItemCode"></param>
        /// <returns></returns>
        private string PS_MM005_GetValidItemCode(string pItemCode)
        {
            string returnValue = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  ISNULL(CONVERT(VARCHAR(10), frozenTo, 120), '') AS [frozenTo], ";
                sQry += "         frozenFor AS [frozenFor]";
                sQry += " FROM    OITM";
                sQry += " WHERE   ItemCode = '" + pItemCode + "'";

                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.Fields.Item("frozenTo").Value == "2999-12-31" && oRecordSet01.Fields.Item("frozenFor").Value == "Y")
                {
                    returnValue = "UnAuthorized"; //미승인
                }
                else if (oRecordSet01.Fields.Item("frozenTo").Value == "2899-12-31" && oRecordSet01.Fields.Item("frozenFor").Value == "Y")
                {
                    returnValue = "UnUsed"; //미사용
                }
                else
                {
                    returnValue = "Authorized"; //승인
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 구매요청 등록
        /// </summary>
        /// <returns></returns>
        private bool PS_MM005_AddPurchaseDemand()
        {
            bool returnValue = false;
            string sQry;
            string Query02;
            string ErrOrdNum = string.Empty;
            string ItemCode;
            string ItemName;
            string BPLId;
            double Qty;
            double Weight;
            string Note;
            string QCYN;
            string UseDept;
            string CntcName;
            string DueDate;
            string DocDate;
            string CntcCode;
            string DeptCode;
            string autoOut;
            string PP030DL;
            string PP030HNo;
            string ProcName;
            string OrdType;
            string OKYN;
            string OkDate;
            string ProcCode;
            string Comments;
            string PP030LNo;
            string OutSize;
            string OutUnit;
            string MaxItemCode = string.Empty;
            string outCode;
            string outNote;
            string MachCode;
            string MachName;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            string ImportYN; //수입품여부
            string EmergYN; //긴급여부
            string errMessage = string.Empty;
            string successMessage = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
                DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();

                if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "60") 
                {
                    sQry = "  Select    right(rtrim('00000' + Convert(Nvarchar(6), Convert(decimal(7,0), Isnull(Max(right(U_ItemCode,6)),'0')) + 1)),6) ";
                    sQry += " From      [@PS_MM005H]";
                    sQry += " Where     U_BPLId = '" + BPLId + "'";
                    sQry += "           and U_OrdType = '" + OrdType + "'";
                    sQry += "           and U_ItemCode <> '1' ";

                    RecordSet01.DoQuery(sQry);

                    MaxItemCode = "90" + BPLId + RecordSet01.Fields.Item(0).Value.ToString().Trim();
                }

                for (int i = 0; i <= oMat01.RowCount - 2; i++)
                {
                    DocDate = oDS_PS_MM005H.GetValue("U_DocDate", i).ToString().Trim();
                    DueDate = oDS_PS_MM005H.GetValue("U_DueDate", i).ToString().Trim();

                    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_PP030DL", i).ToString().Trim()))
                    {
                        PP030DL = "";
                        PP030HNo = "";
                        PP030LNo = "";
                    }
                    else
                    {
                        PP030DL = oDS_PS_MM005H.GetValue("U_PP030DL", i).ToString().Trim();
                        PP030HNo = PP030DL.Split('-')[0];
                        PP030LNo = PP030DL.Split('-')[1];

                        //for (j = 1; j <= PP030DL.Length; j++)
                        //{
                        //    if (Strings.Mid(PP030DL, j, 1) == "-")
                        //    {
                        //        break; // TODO: might not be correct. Was : Exit For
                        //    }
                        //    else
                        //    {
                        //        PP030HNo = PP030HNo + Strings.Mid(PP030DL, j, 1);
                        //    }
                        //}
                        //for (j = 1; j <= Strings.Len(PP030DL); j++)
                        //{
                        //    if (Strings.Mid(PP030DL, j, 1) == "-")
                        //    {
                        //        K = j;
                        //        break; // TODO: might not be correct. Was : Exit For
                        //    }
                        //}
                        //PP030LNo = Strings.Mid(PP030DL, K + 1, Strings.Len(PP030DL) - K);
                    }

                    ItemCode = oDS_PS_MM005H.GetValue("U_ItemCode", i).ToString().Trim();

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "60") //고정자산품의
                    {
                        if (string.IsNullOrEmpty(ItemCode))
                        {
                            ItemCode = "90" + BPLId + codeHelpClass.Right(("00000" + Convert.ToString(Convert.ToInt32(codeHelpClass.Right(MaxItemCode, 6)) + i).TrimStart()).TrimEnd(), 6);
                            //ItemCode = "90" + BPLId + Strings.Right(Strings.RTrim("00000" + Strings.LTrim(Conversion.Str(Conversion.Val(Strings.Right(MaxItemCode, 6)) + i))), 6);
                        }
                    }

                    ItemName = dataHelpClass.Make_ItemName(oDS_PS_MM005H.GetValue("U_ItemName", i).ToString().Trim());
                    
                    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_Qty", i).ToString().Trim()))
                    {
                        Qty = 0;
                    }
                    else
                    {
                        Qty = Convert.ToDouble(oDS_PS_MM005H.GetValue("U_Qty", i).ToString().Trim());
                    }
                    Weight = Convert.ToDouble(oDS_PS_MM005H.GetValue("U_Weight", i).ToString().Trim());
                    UseDept = oDS_PS_MM005H.GetValue("U_UseDept", i).ToString().Trim();
                    autoOut = oDS_PS_MM005H.GetValue("U_Auto", i).ToString().Trim();
                    QCYN = oDS_PS_MM005H.GetValue("U_QCYN", i).ToString().Trim();
                    Note = oDS_PS_MM005H.GetValue("U_Note", i).ToString().Trim();
                    outCode = oDS_PS_MM005H.GetValue("U_OutCode", i).ToString().Trim();
                    outNote = oDS_PS_MM005H.GetValue("U_OutNote", i).ToString().Trim();
                    OKYN = oDS_PS_MM005H.GetValue("U_OKYN", i).ToString().Trim();
                    OkDate = oDS_PS_MM005H.GetValue("U_OKDate", i).ToString().Trim();
                    ProcCode = oDS_PS_MM005H.GetValue("U_ProcCode", i).ToString().Trim();
                    ProcName = oDS_PS_MM005H.GetValue("U_ProcName", i).ToString().Trim();
                    Comments = oDS_PS_MM005H.GetValue("U_Comments", i).ToString().Trim();
                    OutSize = oDS_PS_MM005H.GetValue("U_OutSize", i).ToString().Trim();
                    OutUnit = oDS_PS_MM005H.GetValue("U_OutUnit", i).ToString().Trim();
                    MachCode = oDS_PS_MM005H.GetValue("U_MachCode", i).ToString().Trim(); //설비코드
                    MachName = oDS_PS_MM005H.GetValue("U_MachName", i).ToString().Trim(); //설비명
                    ImportYN = oDS_PS_MM005H.GetValue("U_ImportYN", i).ToString().Trim(); //수입품여부
                    EmergYN = oDS_PS_MM005H.GetValue("U_EmergYN", i).ToString().Trim(); //긴급여부

                    if (!string.IsNullOrEmpty(ItemCode))
                    {
                        sQry = "EXEC [PS_MM005_05] ";
                        sQry += "'" + PP030DL + "',";
                        sQry += "'" + PP030HNo + "',";
                        sQry += "'" + PP030LNo + "',";
                        sQry += "'" + ItemCode + "',";
                        sQry += "'" + ItemName + "',";
                        sQry += "'" + Qty + "',";
                        sQry += "'" + Weight + "',";
                        sQry += "'" + BPLId + "',";
                        sQry += "'" + DocDate + "',";
                        sQry += "'" + DueDate + "',";
                        sQry += "'" + CntcCode + "',";
                        sQry += "'" + CntcName + "',";
                        sQry += "'" + DeptCode + "',";
                        sQry += "'" + UseDept + "',";
                        sQry += "'" + autoOut + "',";
                        sQry += "'" + QCYN + "',";
                        sQry += "'" + Note + "',";
                        sQry += "'" + outCode + "',";
                        sQry += "'" + outNote + "',";
                        sQry += "'" + OKYN + "',";
                        sQry += "'" + OkDate + "',";
                        sQry += "'" + OrdType + "',";
                        sQry += "'" + ProcCode + "',";
                        sQry += "'" + ProcName + "',";
                        sQry += "'" + Comments + "',";
                        sQry += "'" + OutSize + "',";
                        sQry += "'" + OutUnit + "',";
                        sQry += "'" + MachCode + "',";
                        sQry += "'" + MachName + "',";
                        sQry += "'" + ImportYN + "',";
                        sQry += "'" + EmergYN + "',";
                        sQry += "'" + PSH_Globals.oCompany.UserSignature + "'";

                        //선행프로세스 대비 일자체크_S
                        BaseEntry = PP030HNo;
                        BaseLine = "0";
                        DocType = "PS_MM005";
                        CurDocDate = DocDate;

                        Query02 = "EXEC PS_Z_CHECK_DATE '";
                        Query02 += BaseEntry + "','";
                        Query02 += BaseLine + "','";
                        Query02 += DocType + "','";
                        Query02 += CurDocDate + "'";

                        RecordSet02.DoQuery(Query02);

                        if (RecordSet02.Fields.Item("ReturnValue").Value == "True")
                        {
                            RecordSet01.DoQuery(sQry); //등록
                        }
                        else
                        {
                            ErrOrdNum = ErrOrdNum + " [" + ItemCode + "]";
                        }

                        PP030HNo = "";
                        PP030LNo = "";

                        //하나라도 선행프로세스 일자가 빠른 작번이 있으면
                        if (!string.IsNullOrEmpty(ErrOrdNum))
                        {
                            errMessage = "구매요청일은 작업지시일자보다 같거나 늦어야합니다. 확인하십시오." + (char)13 + ErrOrdNum;
                            throw new Exception();
                        }
                        //선행프로세스 대비 일자체크_E
                    }
                }

                successMessage = "구매요청 등록 완료";
              
                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (successMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(successMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }

            return returnValue;
        }

        /// <summary>
        /// 구매요청 수정
        /// </summary>
        /// <returns></returns>
        private bool PS_MM005_UpdatePurchaseDemand()
        {
            bool returnValue = false;
            string sQry;
            string Query02;
            string Query03;
            string ErrOrdNum = string.Empty;
            string ItemCode;
            string ItemName;
            double Qty;
            double Weight;
            string Note;
            string QCYN;
            string UseDept;
            string CntcName;
            string DueDate;
            string DocDate;
            string CntcCode;
            string DeptCode;
            string autoOut;
            string BPLId;
            string PP030HNo;
            string Comments;
            string CgNum;
            string OrdType;
            string OKYN;
            string OkDate;
            string ProcCode;
            string Status;
            string PP030DL;
            string PP030LNo;
            string OutSize;
            string OutUnit;
            string outCode;
            string outNote;
            string MachCode;
            string MachName;
            string ImportYN;
            string EmergYN;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            string errMessage = string.Empty;
            string successMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                for (int i = 0; i <= oMat01.RowCount - 1; i++)
                {
                    Query03 = "select U_UpdtUser from [@PS_MM005H] where U_CgNum ='" + oDS_PS_MM005H.GetValue("U_CgNum", i).ToString().Trim() + "'";
                    RecordSet03.DoQuery(Query03);

                    DocDate = oDS_PS_MM005H.GetValue("U_DocDate", i).ToString().Trim();
                    DueDate = oDS_PS_MM005H.GetValue("U_DueDate", i).ToString().Trim();
                    CgNum = oDS_PS_MM005H.GetValue("U_CgNum", i).ToString().Trim();
                    ItemCode = oDS_PS_MM005H.GetValue("U_ItemCode", i).ToString().Trim();
                    ItemName = dataHelpClass.Make_ItemName(oDS_PS_MM005H.GetValue("U_ItemName", i).ToString().Trim());
                    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_Qty", i).ToString().Trim()))
                    {
                        Qty = 0;
                    }
                    else
                    {
                        Qty = Convert.ToDouble(oDS_PS_MM005H.GetValue("U_Qty", i).ToString().Trim());
                    }
                    Weight = Convert.ToDouble(oDS_PS_MM005H.GetValue("U_Weight", i).ToString().Trim());
                    UseDept = oDS_PS_MM005H.GetValue("U_UseDept", i).ToString().Trim();
                    autoOut = oDS_PS_MM005H.GetValue("U_Auto", i).ToString().Trim();
                    QCYN = oDS_PS_MM005H.GetValue("U_QCYN", i).ToString().Trim();
                    Note = oDS_PS_MM005H.GetValue("U_Note", i).ToString().Trim();
                    outCode = oDS_PS_MM005H.GetValue("U_OutCode", i).ToString().Trim();
                    outNote = oDS_PS_MM005H.GetValue("U_OutNote", i).ToString().Trim();
                    OKYN = oDS_PS_MM005H.GetValue("U_OKYN", i).ToString().Trim();
                    OkDate = oDS_PS_MM005H.GetValue("U_OKDate", i).ToString().Trim();
                    ProcCode = oDS_PS_MM005H.GetValue("U_ProcCode", i).ToString().Trim();
                    Comments = oDS_PS_MM005H.GetValue("U_Comments", i).ToString().Trim();
                    Status = oDS_PS_MM005H.GetValue("U_Status", i).ToString().Trim();
                    OrdType = oDS_PS_MM005H.GetValue("U_OrdType", i).ToString().Trim();
                    BPLId = oDS_PS_MM005H.GetValue("U_BPLId", i).ToString().Trim();
                    CntcCode = oDS_PS_MM005H.GetValue("U_CntcCode", i).ToString().Trim();
                    CntcName = oDS_PS_MM005H.GetValue("U_CntcName", i).ToString().Trim();
                    DeptCode = oDS_PS_MM005H.GetValue("U_DeptCode", i).ToString().Trim();
                    OutSize = oDS_PS_MM005H.GetValue("U_OutSize", i).ToString().Trim();
                    OutUnit = oDS_PS_MM005H.GetValue("U_OutUnit", i).ToString().Trim();
                    MachCode = oDS_PS_MM005H.GetValue("U_MachCode", i).ToString().Trim(); //설비코드
                    MachName = oDS_PS_MM005H.GetValue("U_MachName", i).ToString().Trim(); //설비명
                    ImportYN = oDS_PS_MM005H.GetValue("U_ImportYN", i).ToString().Trim(); //수입품여부
                    EmergYN = oDS_PS_MM005H.GetValue("U_EmergYN", i).ToString().Trim(); //긴급여부

                    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_PP030DL", i).ToString().Trim()))
                    {
                        PP030DL = "";
                        PP030HNo = "";
                        PP030LNo = "";
                    }
                    else
                    {
                        PP030DL = oDS_PS_MM005H.GetValue("U_PP030DL", i).ToString().Trim();
                        PP030HNo = PP030DL.Split('-')[0];
                        PP030LNo = PP030DL.Split('-')[1];
                        //for (j = 1; j <= Strings.Len(PP030DL); j++)
                        //{
                        //    if (Strings.Mid(PP030DL, j, 1) == "-")
                        //    {
                        //        break; // TODO: might not be correct. Was : Exit For
                        //    }
                        //    else
                        //    {
                        //        PP030HNo = PP030HNo + Strings.Mid(PP030DL, j, 1);
                        //    }
                        //}
                        //for (j = 1; j <= Strings.Len(PP030DL); j++)
                        //{
                        //    if (Strings.Mid(PP030DL, j, 1) == "-")
                        //    {
                        //        K = j;
                        //        break; // TODO: might not be correct. Was : Exit For
                        //    }
                        //}
                        //PP030LNo = Strings.Mid(PP030DL, K + 1, Strings.Len(PP030DL) - K);
                    }

                    sQry = "UPDATE [@PS_MM005H] ";
                    sQry += "SET ";
                    sQry += "U_PP030DL = '" + PP030DL + "', ";
                    sQry += "U_PP030HNo = '" + PP030HNo + "', ";
                    sQry += "U_PP030LNo = '" + PP030LNo + "', ";
                    sQry += "U_ItemCode = '" + ItemCode + "', ";
                    sQry += "U_ItemName = '" + ItemName + "', ";
                    sQry += "U_Qty = '" + Qty + "', ";
                    sQry += "U_Weight = '" + Weight + "', ";
                    sQry += "U_CgNum = '" + CgNum + "', ";
                    sQry += "U_DocDate = '" + DocDate + "', ";
                    sQry += "U_DueDate = '" + DueDate + "', ";
                    sQry += "U_CntcCode = '" + CntcCode + "', ";
                    sQry += "U_CntcName = '" + CntcName + "', ";
                    sQry += "U_DeptCode = '" + DeptCode + "', ";
                    sQry += "U_UseDept = '" + UseDept + "', ";
                    sQry += "U_Auto = '" + autoOut + "', ";
                    sQry += "U_QCYN = '" + QCYN + "', ";
                    sQry += "U_Note = '" + Note + "', ";
                    sQry += "U_OutCode = '" + outCode + "', ";
                    sQry += "U_OutNote = '" + outNote + "', ";
                    sQry += "U_OKYN = '" + OKYN + "', ";
                    if (string.IsNullOrEmpty(OkDate))
                    {
                        sQry += "U_OKDate = NULL, ";
                    }
                    else
                    {
                        sQry += "U_OKDate = '" + OkDate + "', ";
                    }
                    sQry += "U_OrdType = '" + OrdType + "', ";
                    sQry += "U_ProcCode = '" + ProcCode + "', ";
                    sQry += "U_Comments = '" + Comments + "', ";
                    sQry += "U_OutSize = '" + OutSize + "', ";
                    sQry += "U_OutUnit = '" + OutUnit + "', ";
                    sQry += "U_Status = '" + Status + "', ";
                    sQry += "U_MachCode = '" + MachCode + "', ";
                    sQry += "U_MachName = '" + MachName + "', ";
                    sQry += "U_ImportYN = '" + ImportYN + "', ";
                    sQry += "U_EmergYN = '" + EmergYN + "' ";

                    if (oDS_PS_MM005H.GetValue("U_UpdtUser", i).ToString().Trim() != RecordSet03.Fields.Item(0).Value.ToString().Trim())
                    {
                        sQry += ", UpdateDate = GETDATE(), ";
                        sQry += "U_UpdtUser = '" + PSH_Globals.oCompany.UserSignature + "' ";
                    }

                    sQry += "Where DocEntry = '" + CgNum + "' ";

                    //선행프로세스 대비 일자체크_S
                    BaseEntry = PP030HNo;
                    BaseLine = "0";
                    DocType = "PS_MM005";
                    CurDocDate = DocDate;

                    Query02 = "EXEC PS_Z_CHECK_DATE '";
                    Query02 += BaseEntry + "','";
                    Query02 += BaseLine + "','";
                    Query02 += DocType + "','";
                    Query02 += CurDocDate + "'";

                    RecordSet02.DoQuery(Query02);
                    //선행프로세스 대비 일자체크_E

                    if (RecordSet02.Fields.Item("ReturnValue").Value == "True")
                    {
                        RecordSet01.DoQuery(sQry);
                    }
                    else
                    {
                        ErrOrdNum = ErrOrdNum + " [" + ItemCode + "]";
                    }

                    //하나라도 선행프로세스 일자가 빠른 작번이 있으면
                    if (!string.IsNullOrEmpty(ErrOrdNum))
                    {
                        errMessage = "구매요청일은 작업지시일자보다 같거나 늦어야합니다. 확인하십시오." + (char)13 + ErrOrdNum;
                        throw new Exception();
                    }
                }

                successMessage = "구매요청 수정 완료";

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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (successMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(successMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet03);
            }

            return returnValue;
        }

        /// <summary>
        /// AP대변 메모 입력건에 대한 처리(U_Status를 "C"로 변경, 이력 저장)
        /// 1. 해당 청구건에 대한 AP대변메모 등록 여부 확인
        /// 2. 구매요청의 U_Status 변경("O" -> "C")
        /// 3. 변경 이력 저장
        /// </summary>
        private void PS_MM005_UpdateAPCreditMemoData()
        {
            string sQry;
            string errMessage = string.Empty;
            string successMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //AP대변메모 등록 여부 확인
                sQry = "EXEC PS_MM005_06 '";
                sQry += oLast_RightClick_CgNum + "'"; //구매요청 문서번호
                oRecordSet01.DoQuery(sQry);

                //AP대변메모 등록 안된 경우
                if (string.IsNullOrEmpty(oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim()))
                {
                    errMessage = "AP대변메모가 등록되지않은 구매요청입니다. 확인하세요.";
                    throw new Exception();
                }

                //U_Status 변경
                sQry = "EXEC PS_MM005_07 '";
                sQry += oLast_RightClick_CgNum + "'"; //구매요청 문서번호
                oRecordSet01.DoQuery(sQry);

                //변경 이력 저장
                sQry = "EXEC PS_MM005_08 '";
                sQry += oLast_RightClick_CgNum + "','"; //구매요청 문서번호
                sQry += PSH_Globals.oCompany.UserName + "'"; //UserID
                oRecordSet01.DoQuery(sQry);

                successMessage = "처리 완료";
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (successMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(successMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_MM005_DelHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("OrdType").Specific.Value.ToString().Trim()))
                {
                    errMessage = "품목구분은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                    {
                        errMessage = "청구인은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim())) 
                    {
                        errMessage = "청구부서는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                }

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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            
            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_MM005_DelMatrixSpaceLine()
        {
            bool returnValue = false;
            int j;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oMat01.FlushToDataSource();

                //라인
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_ItemCode", 0).ToString().Trim()) && string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_ItemName", 0).ToString().Trim()))
                {
                    j = 0;
                }
                else
                {
                    j = oMat01.VisualRowCount;
                }

                for (int i = 0; i <= j - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_ItemCode", i)) && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        errMessage = (i + 1) + "번 라인의 품목코드가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oDS_PS_MM005H.GetValue("U_Weight", i)) == 0)
                    {
                        errMessage = (i + 1) + "번 라인의 수량/중량이 0보다 작거나 같을 수 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_DocDate", i)))
                    {
                        errMessage = (i + 1) + "번 라인의 청구일자가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_DueDate", i)))
                    {
                        errMessage = (i + 1) + "번 라인의 납기일자가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "30" || oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "40") //외주제작, 외주가공 품의인 경우 
                    {
                        if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_OutCode", i))) //외주사유코드 필수 입력
                        {
                            errMessage = "외주제작품의, 가공비품의인 경우는 외주사유가 필수입니다. " + (i + 1) + "번 라인의 [외주사유코드]가 없습니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_OutNote", i))) //외주사유내용 필수 입력
                        {
                            errMessage = "외주제작품의, 가공비품의인 경우는 외주사유가 필수입니다. " + (i + 1) + "번 라인의 [외주사유내용]이 없습니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
                    else if (Convert.ToInt32(oDS_PS_MM005H.GetValue("U_DocDate", i)) >= Convert.ToInt32(oDS_PS_MM005H.GetValue("U_DueDate", i)))
                    {
                        errMessage = "납기일자가 청구일자와 같거나 이전일 입니다. " + (i + 1) + "번 라인의 납기일을 확인하세요.";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "40" && string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_OutUnit", i))) //외주제작 품의인 경우 단위 필수 입력
                    {
                        errMessage = "외주제작품의인 경우는 단위가 필수입니다. " + (i + 1) + "번 라인의 [단위]가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (PS_MM005_GetValidItemCode(oDS_PS_MM005H.GetValue("U_ItemCode", i).ToString().Trim()) == "UnAuthorized") //미승인
                    {
                        errMessage = (i + 1) + "번 라인의 품목은 미승인 품목입니다. 승인 후 구매요청을 진행하십시오.";
                        throw new Exception();
                    }
                    else if (PS_MM005_GetValidItemCode(oDS_PS_MM005H.GetValue("U_ItemCode", i).ToString().Trim()) == "UnUsed")
                    {
                        errMessage = (i + 1) + "번 라인의 품목은 비활성 품목입니다.";
                        throw new Exception();
                    }
                    else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oDS_PS_MM005H.GetValue("U_DocDate", i), oForm.TypeEx) == false) //마감상태 체크
                    {
                        errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. " + (i + 1) + "번 라인의 청구일자를 확인하고, 회계부서로 문의하세요.";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 조회전 모드 변경시 조회조건 저장
        /// </summary>
        private void PS_MM005_SetPreSearchData()
        {
            try
            {
                searchData.SetOrdType(oForm.Items.Item("OrdType").Specific.Value);
                searchData.SetBPLID(oForm.Items.Item("BPLId").Specific.Value);
                searchData.SetCntcCode(oForm.Items.Item("CntcCode").Specific.Value);
                searchData.SetCntcName(oForm.Items.Item("CntcName").Specific.Value);
                searchData.SetDeptCode(oForm.Items.Item("DeptCode").Specific.Value);
                searchData.SetDocDateFr(oForm.Items.Item("DocDateFr").Specific.Value);
                searchData.SetDocDateTo(oForm.Items.Item("DocDateTo").Specific.Value);
                searchData.SetCgNumFr(oForm.Items.Item("CgNumFr").Specific.Value);
                searchData.SetCgNumTo(oForm.Items.Item("CgNumTo").Specific.Value);
                searchData.SetItemCode(oForm.Items.Item("ItemCode").Specific.Value);
                searchData.SetItmBSort(oForm.Items.Item("ItmBSort").Specific.Value);
                searchData.SetItmMSort(oForm.Items.Item("ItmMSort").Specific.Value);
                searchData.SetItemType(oForm.Items.Item("ItemType").Specific.Value);
                searchData.SetOKYN(oForm.Items.Item("OKYN").Specific.Value);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 조회후 모드 변경시 조회전 조건 값 할당
        /// </summary>
        private void PS_MM005_GetPreSearchData()
        {
            try
            {
                oForm.Freeze(true);
                oForm.DataSources.UserDataSources.Item("OrdType").Value = searchData.GetOrdType();
                oForm.DataSources.UserDataSources.Item("BPLId").Value = searchData.GetBPLID();
                oForm.DataSources.UserDataSources.Item("CntcCode").Value = searchData.GetCntcCode();
                oForm.DataSources.UserDataSources.Item("CntcName").Value = searchData.GetCntcName();
                oForm.DataSources.UserDataSources.Item("DeptCode").Value = searchData.GetDeptCode();
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = searchData.GetDocDateFr();
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = searchData.GetDocDateTo();
                oForm.DataSources.UserDataSources.Item("CgNumFr").Value = searchData.GetCgNumFr();
                oForm.DataSources.UserDataSources.Item("CgNumTo").Value = searchData.GetCgNumTo();
                oForm.DataSources.UserDataSources.Item("ItemCode").Value = searchData.GetItemCode();
                oForm.DataSources.UserDataSources.Item("ItmBSort").Value = searchData.GetItmBSort();
                oForm.DataSources.UserDataSources.Item("ItmMSort").Value = searchData.GetItmMSort();
                oForm.DataSources.UserDataSources.Item("ItemType").Value = searchData.GetItemType();
                oForm.DataSources.UserDataSources.Item("OKYN").Value = searchData.GetOKYN();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM005_DelHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM005_DelMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_MM005_AddPurchaseDemand() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_MM005_AddMatrixRow(0, true);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM005_DelMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_MM005_UpdatePurchaseDemand() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PS_MM005_SetPreSearchData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_MM005_GetPreSearchData();
                            PS_MM005_LoadCaption();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Close();
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (PS_MM005_DelHeaderSpaceLine() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        PS_MM005_LoadData();
                        PS_MM005_SetPreSearchData();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PS_MM005_GetPreSearchData();
                        PS_MM005_LoadCaption();
                        PS_MM005_EnableFormItem();
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        if (PS_MM005_SetMatrixData() == false)
                        {
                            BubbleEvent = false;
                            return;
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
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
                            {
                                PS_SM010 tempForm = new PS_SM010();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                if (oForm.Items.Item("OrdType").Specific.Value == "60")
                                {
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) 
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                        BubbleEvent = false;
                                    }
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) 
                                    {
                                        if (oForm.Items.Item("BOM_CHECK").Specific.Checked == true)
                                        {
                                            PS_SM030 tempForm = new PS_SM030();
                                            tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, "");
                                            BubbleEvent = false;
                                        }
                                        else
                                        {
                                            PS_SM010 tempForm = new PS_SM010();
                                            tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                            BubbleEvent = false;
                                        }
                                    }
                                }
                            }
                            else if (pVal.ColUID == "PP030DL")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030DL").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) 
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "CntcCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "MachCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("MachCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "ProcCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ProcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "OrdType" || pVal.ItemUID == "BPLId")
                    {
                        oMat01.Clear();
                        oDS_PS_MM005H.Clear();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM005_AddMatrixRow(0, false);
                        }

                        if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "30" && oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "2") 
                        {
                            oMat01.Columns.Item("PP030DL").Editable = true;
                            oMat01.Columns.Item("ItemCode").Editable = false;
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                        }
                        else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "40" && oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "2") 
                        {
                            oMat01.Columns.Item("PP030DL").Editable = true;
                            oMat01.Columns.Item("ItemCode").Editable = false;
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                        }
                        else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "60") 
                        {
                            oMat01.Columns.Item("ItemCode").Editable = true;
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                            oMat01.Columns.Item("Qty").Editable = false;
                        }
                        else
                        {
                            oMat01.Columns.Item("PP030DL").Editable = false;
                            oMat01.Columns.Item("ItemCode").Editable = true;
                            oMat01.Columns.Item("ItemName").Editable = false;
                            oMat01.Columns.Item("OutSize").Editable = false;
                            oMat01.Columns.Item("OutUnit").Editable = false;
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        oMat01.Columns.Item("UpdtUser").Cells.Item(pVal.Row).Specific.Value = PSH_Globals.oCompany.UserSignature;
                        if (pVal.ColUID == "OKYN")
                        {
                            oMat01.FlushToDataSource();
                            if (oMat01.Columns.Item("OKYN").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "Y") 
                            {
                                oMat01.Columns.Item("OKDate").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            }
                            else if (oMat01.Columns.Item("OKYN").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "N") 
                            {
                                oMat01.Columns.Item("OKDate").Cells.Item(pVal.Row).Specific.Value = "";
                            }
                        }
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else
                        {
                            PS_MM005_SetPreSearchData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_MM005_GetPreSearchData();
                            PS_MM005_LoadCaption();
                        }
                    }
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    //if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "OKYN" && oMenuUID == "PS_MM005_1")
                    //{
                    //    string OKYN = string.Empty;

                    //    oForm.Freeze(true);
                    //    oMat01.FlushToDataSource();
                    //    if (string.IsNullOrEmpty(oDS_PS_MM005H.GetValue("U_OKYN", 0).ToString().Trim()) || oDS_PS_MM005H.GetValue("U_OKYN", 0).ToString().Trim() == "N")
                    //    {
                    //        OKYN = "Y";
                    //    }
                    //    else if (oDS_PS_MM005H.GetValue("U_OKYN", 0).ToString().Trim() == "Y")
                    //    {
                    //        OKYN = "N";
                    //    }
                    //    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    //    {
                    //        oDS_PS_MM005H.SetValue("U_OKYN", i, OKYN);
                    //        if (OKYN == "Y")
                    //        {
                    //            oDS_PS_MM005H.SetValue("U_OKDate", i, DateTime.Now.ToString("yyyyMMdd"));
                    //        }
                    //        else if (OKYN == "N")
                    //        {
                    //            oDS_PS_MM005H.SetValue("U_OKDate", i, "");
                    //        }
                    //    }
                    //    oMat01.LoadFromDataSource();
                    //    oMat01.Columns.Item("OKDate").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //    {
                    //    }
                    //    else
                    //    {
                    //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    //        PS_MM005_LoadCaption();
                    //    }
                    //    oForm.Freeze(false);
                    //}
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            PS_MM005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ItemName")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "PP030DL")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "Weight")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "CntcCode")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ProcCode")
                            {
                                PS_MM005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                            }
                            else
                            {
                                PS_MM005_SetPreSearchData();
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                PS_MM005_GetPreSearchData();
                                PS_MM005_LoadCaption();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM005H);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                sQry = " Select U_OKYN From [@PS_MM005H] ";
                                sQry += "where U_CgNum = '" + oLast_RightClick_CgNum + "'";
                                oRecordSet01.DoQuery(sQry);
                                if (oRecordSet01.Fields.Item(0).Value == "Y")
                                {
                                    PSH_Globals.SBO_Application.MessageBox("결재가 된 구매요청은 삭제할 수 없습니다.");
                                    BubbleEvent = false;
                                    return;
                                }

                                sQry = " Select Count(a.U_CGNo) From [@PS_MM010L] a Inner Join [@PS_MM010H] b On a.DocEntry = b.DocEntry ";
                                sQry += "where a.U_CGNo = '" + oLast_RightClick_CgNum + "' And b.Status = 'O'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item(0).Value == 0)
                                {
                                    if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 구매요청을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소") == 1)
                                    {
                                        sQry = "Delete [@PS_MM005H] Where DocEntry = '" + oLast_RightClick_CgNum + "'";
                                        oRecordSet01.DoQuery(sQry);
                                        oLast_RightClick_CgNum = 0;
                                        PSH_Globals.SBO_Application.StatusBar.SetText("처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    }
                                    else
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.MessageBox("해당 구매요청건에 해당하는 구매견적이 있습니다. 삭제할 수 없습니다.");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            break;
                        case "1299": //행닫기(취소)
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                sQry = " Select U_OKYN From [@PS_MM005H] ";
                                sQry += "where U_CgNum = '" + oLast_RightClick_CgNum + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item(0).Value == "Y")
                                {
                                    PSH_Globals.SBO_Application.MessageBox("결재처리된 구매요청은 닫기(취소)할 수 없습니다.");
                                    BubbleEvent = false;
                                    return;
                                }

                                sQry = " Select Count(a.U_CGNo) From [@PS_MM010L] a Inner Join [@PS_MM010H] b On a.DocEntry = b.DocEntry ";
                                sQry += "where a.U_CGNo = '" + oLast_RightClick_CgNum + "' And b.Status = 'O'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item(0).Value == 0)
                                {
                                    if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 구매요청을 닫기(취소)합니다. 진행하시겠습니까?", 1, "&확인", "&취소") == 1)
                                    {
                                        sQry = "  UPDATE [@PS_MM005H] ";
                                        sQry += " SET    Status = 'C',";
                                        sQry += "        Canceled = 'Y',";
                                        sQry += "        UpdateDate = GETDATE(),";
                                        sQry += "        UserSign = '" + PSH_Globals.oCompany.UserSignature + "'";
                                        sQry += " Where  DocEntry = '" + oLast_RightClick_CgNum + "'";
                                        oRecordSet01.DoQuery(sQry);
                                        oLast_RightClick_CgNum = 0;
                                        PSH_Globals.SBO_Application.StatusBar.SetText("처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    }
                                    else
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.MessageBox("해당 구매요청건에 해당하는 구매견적이 있습니다. 닫기(취소)할 수 없습니다.");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "APCreditMemo": //[AP대변 메모] 입력 확인
                            if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 구매요청을 AP대변메모 등록으로 처리합니다. 처리 후 복원할 수 없습니다. 진행하시겠습니까?", 1, "&확인", "&취소") == 1)
                            {
                                PS_MM005_UpdateAPCreditMemoData();
                                oLast_RightClick_CgNum = 0;
                            }
                            else
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_MM005H.RemoveRecord(oDS_PS_MM005H.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                                oForm.Freeze(false);
                            }
                            break;
                        case "1281": //찾기
                            PS_MM005_SetPreSearchData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_MM005_GetPreSearchData();
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_MM005_AddMatrixRow(0, true);
                            PS_MM005_EnableFormItem();
                            PS_MM005_LoadCaption();
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_MM005_SetPreSearchData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_MM005_GetPreSearchData();
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_MM005_AddMatrixRow(0, true);
                            PS_MM005_EnableFormItem();
                            PS_MM005_LoadCaption();
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_MM005_EnableFormItem();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.Value.ToString().Trim()))
                                {
                                    if (oDS_PS_MM005H.GetValue("Status", 0) == "O")
                                    {
                                        PS_MM005_AddMatrixRow(oMat01.RowCount, false);
                                    }
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.Menus oMenus = null;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row > 0 && pVal.Row <= oMat01.RowCount)
                    {
                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("CgNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) 
                        {
                            oLast_RightClick_CgNum = Convert.ToInt32(oMat01.Columns.Item("CgNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //청구번호 저장
                        }

                        ////oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                        ////oCreationPackage = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                        //oCreationPackage = ((SAPbouiCOM.MenuCreationParams)PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams));
                        //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        //oCreationPackage.UniqueID = "APCreditMemo";
                        //oCreationPackage.String = "[AP대변 메모] 입력 확인";
                        //oCreationPackage.Enabled = true;
                        //oCreationPackage.Position = -1;

                        //oMenuItem = PSH_Globals.SBO_Application.Menus.Item("1280");
                        //oMenus = oMenuItem.SubMenus;
                        //oMenus.AddEx(oCreationPackage);
                        ////PSH_Globals.SBO_Application.Menus.Item("1280").SubMenus.AddEx(oCreationPackage);



                        //oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        //oCreationPackage.UniqueID = "OnlyOnRC";
                        //oCreationPackage.String = "Only On Right Click";
                        //oCreationPackage.Enabled = true;

                        //oMenuItem = PSH_Globals.SBO_Application.Menus.Item("1280"); // Data'
                        //oMenus = oMenuItem.SubMenus;
                        //oMenus.AddEx(oCreationPackage);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    //"[AP대변 메모] 입력 확인" 삭제
                    if (pVal.ItemUID == "Mat01" && pVal.Row > 0)
                    {
                        if (oMat01.RowCount >= pVal.Row)
                        {
                            //PSH_Globals.SBO_Application.Menus.RemoveEx("OnlyOnRC");
                            //oMenus.RemoveEx("APCreditMemo");
                            //PSH_Globals.SBO_Application.Menus.RemoveEx("APCreditMemo");
                        }
                    }
                }
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oCreationPackage != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage);
                }

                if (oMenuItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem);
                }

                if (oMenus != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenus);
                }
            }
        }
    }
}
