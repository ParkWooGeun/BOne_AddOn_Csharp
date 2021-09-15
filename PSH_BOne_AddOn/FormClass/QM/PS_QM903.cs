using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외부실패비용 등록
	/// </summary>
	internal class PS_QM903 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_QM903H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM903L; //등록라인
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM903.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_QM903_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_QM903");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
				PS_QM903_CreateItems();
				PS_QM903_SetComboBox();
				PS_QM903_SetDocEntry();
				//PS_QM903_AddMatrixRow(1, 0, true);
				//PS_QM903_EnableFormItem();

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", true); //행삭제
				oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_QM903_CreateItems()
        {
            try
            {
                oDS_PS_QM903H = oForm.DataSources.DBDataSources.Item("@PS_QM903H");
                oDS_PS_QM903L = oForm.DataSources.DBDataSources.Item("@PS_QM903L");

                oMat01 = oForm.Items.Item("Mat01").Specific;

                oDS_PS_QM903H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyDDdd"));
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보에 기본값설정
        /// </summary>
        private void PS_QM903_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("FType").Specific.ValidValues.Add("10", "실패비용");
                oForm.Items.Item("FType").Specific.ValidValues.Add("20", "출장처리비용");
                oForm.Items.Item("FType").Specific.ValidValues.Add("30", "고객크레임 처리비용");
                oForm.Items.Item("FType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_QM903_SetDocEntry()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM903'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
















        #region PS_QM903_EnableFormItem
        //private void PS_QM903_EnableFormItem()
        //{
        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //    {
        //        oForm.Items.Item("DocEntry").Enabled = true;

        //    }
        //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //    {
        //        oForm.Items.Item("DocEntry").Enabled = false;

        //    }
        //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //    {
        //        oForm.Items.Item("DocEntry").Enabled = false;

        //    }

        //    return;
        //PS_QM903_EnableFormItem_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "PS_QM903_EnableFormItem_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion



        #region PS_QM903_AddMatrixRow
        ////*******************************************************************
        ////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
        ////*******************************************************************
        //private void PS_QM903_AddMatrixRow(short oMat, int oRow, ref bool Insert_YN = false)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement


        //    switch (oMat)
        //    {
        //        case 1:
        //            //oMat01
        //            if (Insert_YN == false)
        //            {
        //                oRow = oMat01.RowCount;
        //                oDS_PS_QM903L.InsertRecord((oRow));
        //            }
        //            //수입내역
        //            oDS_PS_QM903L.Offset = oRow;
        //            oDS_PS_QM903L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //            //            oDS_PS_QM903L.setValue "U_ItmBsort", oRow, ""
        //            //            oDS_PS_QM903L.setValue "U_FAmt01", oRow, ""
        //            //            oDS_PS_QM903L.setValue "U_FAmt02", oRow, ""
        //            //            oDS_PS_QM903L.setValue "U_TotFAmt", oRow, ""
        //            oMat01.LoadFromDataSource();
        //            break;

        //            //        Case 2: 'oMat02
        //            //            If Insert_YN = False Then
        //            //                oRow = oMat2.RowCount
        //            //                oDS_ZPP140M.InsertRecord (oRow)
        //            //            End If
        //            //            '수출내역
        //            //            oDS_ZPP140M.Offset = oRow
        //            //            oDS_ZPP140M.setValue "LineId", oRow, oRow + 1
        //            //            oDS_ZPP140M.setValue "U_ConfDate", oRow, ""
        //            //            oDS_ZPP140M.setValue "U_ConfNo", oRow, ""
        //            //            oDS_ZPP140M.setValue "U_Size", oRow, ""
        //            //            oDS_ZPP140M.setValue "U_ExpQty", oRow, ""
        //            //            oDS_ZPP140M.setValue "U_RfndQty", oRow, ""
        //            //            oMat02.LoadFromDataSource

        //    }
        //    return;
        //PS_QM903_AddMatrixRow_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "PS_QM903_AddMatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_QM903_FlushToItemValue
        //private void PS_QM903_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int j = 0;
        //    int i = 0;
        //    int cnt = 0;
        //    string DocNum = null;
        //    string LineId = null;
        //    short ErrNum = 0;
        //    string sQry = null;
        //    decimal FAmt01 = default(decimal);
        //    decimal FAmt02 = default(decimal);
        //    decimal TotFAmt = default(decimal);

        //    SAPbobsCOM.Recordset oRecordSet = null;

        //    oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    //--------------------------------------------------------------
        //    //Header--------------------------------------------------------
        //    switch (oUID)
        //    {
        //        case "CardCode":
        //            ////거래처명 검색
        //            sQry = "select CardName from OCRD where CardCode = '" + Strings.Trim(oDS_PS_QM903H.GetValue("U_CardCode", 0)) + "'";
        //            oRecordSet.DoQuery(sQry);
        //            oDS_PS_QM903H.SetValue("U_CardName", 0, Strings.Trim(oRecordSet.Fields.Item(0).Value));
        //            break;
        //    }

        //    //--------------------------------------------------------------
        //    //Line----------------------------------------------------------
        //    if (oUID == "Mat01")
        //    {
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oDS_PS_QM903L.SetValue("U_" + oCol, oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value);
        //        switch (oCol)
        //        {
        //            case "ItmBsort":

        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "select Name from [@PSH_ITMBSORT] where Code = '" + Strings.Trim(oMat01.Columns.Item("ItmBsort").Cells.Item(oRow).Specific.Value) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                oDS_PS_QM903L.SetValue("U_ItmBname", oRow - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));

        //                if (oRow == oMat01.RowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_QM903L.GetValue("U_ItmBsort", oRow - 1))))
        //                {
        //                    //// 다음 라인 추가
        //                    PS_QM903_AddMatrixRow(1, 0, ref false);
        //                    oMat01.Columns.Item("ItmMsort").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                }
        //                break;

        //            case "ItmMsort":
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "select U_CodeName from [@PSH_ITMMSORT] where U_Code = '" + Strings.Trim(oMat01.Columns.Item("ItmMsort").Cells.Item(oRow).Specific.Value) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                oDS_PS_QM903L.SetValue("U_ItmMname", oRow - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
        //                break;


        //            //                oMat01.FlushToDataSource
        //            //                oForm.Freeze False

        //            //                oDS_PS_QM903L.Offset = oRow - 1
        //            //                'oMat01.SetLineData oRow
        //            //
        //            //                '--------------------------------------------------------------------------------------------
        //            case "CntcCode":
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + Strings.Trim(oMat01.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                oDS_PS_QM903L.SetValue("U_CntcName", oRow - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
        //                break;

        //            case "OrdNum":
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "Select U_ItemName From [@PS_PP030H] Where U_ItemCode = '" + Strings.Trim(oMat01.Columns.Item("OrdNum").Cells.Item(oRow).Specific.Value) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                oDS_PS_QM903L.SetValue("U_OrdName", oRow - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
        //                break;
        //            case "ItemCode":
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "Select ItemName From [OITM] Where ItemCode = '" + Strings.Trim(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                oDS_PS_QM903L.SetValue("U_ItemName", oRow - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
        //                break;

        //        }

        //        oMat01.LoadFromDataSource();
        //        oForm.Freeze(false);
        //        oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //    }

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    return;
        //PS_QM903_FlushToItemValue_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "PS_QM903_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_QM903_DelHeaderSpaceLine
        //private bool PS_QM903_DelHeaderSpaceLine()
        //{
        //    bool returnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short ErrNum = 0;

        //    ErrNum = 0;

        //    //// Check
        //    switch (true)
        //    {
        //        case string.IsNullOrEmpty(Strings.Trim(oDS_PS_QM903H.GetValue("U_DocDate", 0))):
        //            ErrNum = 1;
        //            goto PS_QM903_DelHeaderSpaceLine_Error;
        //            break;
        //    }

        //    returnValue = true;
        //    return returnValue;
        //PS_QM903_DelHeaderSpaceLine_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "등록일자는 필수사항입니다. 확인하여 주십시오.", ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "PS_QM903_DelHeaderSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    returnValue = false;
        //    return returnValue;
        //}
        #endregion

        #region PS_QM903_DelMatrixSpaceLine
        //private bool PS_QM903_DelMatrixSpaceLine()
        //{
        //    bool returnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    short ErrNum = 0;
        //    SAPbobsCOM.Recordset oRecordSet = null;
        //    string sQry = null;

        //    oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    ErrNum = 0;

        //    oMat01.FlushToDataSource();

        //    //// 라인
        //    //// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
        //    if (oMat01.VisualRowCount == 1)
        //    {
        //        ErrNum = 1;
        //        goto PS_QM903_DelMatrixSpaceLine_Error;
        //    }

        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    ////마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
        //    ////oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    if (oMat01.VisualRowCount > 0)
        //    {
        //        //// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
        //        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
        //        {
        //            oDS_PS_QM903L.Offset = i;
        //            if (string.IsNullOrEmpty(Strings.Trim(oDS_PS_QM903L.GetValue("U_ItmBsort", i))))
        //            {
        //                ErrNum = 2;
        //                oMat01.Columns.Item("ItmBsort").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                goto PS_QM903_DelMatrixSpaceLine_Error;
        //            }
        //        }
        //    }

        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    ////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
        //    ////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    if (oMat01.VisualRowCount > 0)
        //    {
        //        oDS_PS_QM903L.RemoveRecord(oDS_PS_QM903L.Size - 1);
        //        //// Mat1에 마지막라인(빈라인) 삭제
        //    }
        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    //행을 삭제하였으니 DB데이터 소스를 다시 가져온다
        //    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //    oMat01.LoadFromDataSource();

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    returnValue = true;
        //    return returnValue;
        //PS_QM903_DelMatrixSpaceLine_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하여 주십시오.", ref "E");
        //    }
        //    else if (ErrNum == 2)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "구분(대분류)는 필수사항입니다. 확인하여 주십시오.", ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "PS_QM903_DelMatrixSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    returnValue = false;
        //    return returnValue;
        //}
        #endregion





        #region Raise_ItemEvent
        ////****************************************************************************************************************
        ////// ItemEventHander
        ////****************************************************************************************************************
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    int ErrNum = 0;
        //    SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //    object ChildForm01 = null;
        //    ChildForm01 = new PS_SM010();

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "1")
        //                {
        //                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //                    {
        //                        if (PS_QM903_DelHeaderSpaceLine() == false)
        //                        {
        //                            BubbleEvent = false;
        //                            return;
        //                        }
        //                        if (PS_QM903_DelMatrixSpaceLine() == false)
        //                        {
        //                            BubbleEvent = false;
        //                            return;
        //                        }

        //                    }
        //                }
        //                break;

        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                if (pVal.CharPressed == 9)
        //                {
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                    ////라인
        //                    if (pVal.ItemUID == "Mat01")
        //                    {
        //                        if (pVal.ColUID == "ItmBsort")
        //                        {
        //                            //UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItmBsort").Cells.Item(pVal.Row).Specific.Value))
        //                            {
        //                                SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                                BubbleEvent = false;
        //                            }
        //                        }
        //                        if (pVal.ColUID == "ItmMsort")
        //                        {
        //                            //UPGRADE_WARNING: oMat01.Columns(ItmMsort).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItmMsort").Cells.Item(pVal.Row).Specific.Value))
        //                            {
        //                                SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                                BubbleEvent = false;
        //                            }
        //                        }
        //                        if (pVal.ColUID == "CntcCode")
        //                        {
        //                            //UPGRADE_WARNING: oMat01.Columns(CntcCode).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            if (string.IsNullOrEmpty(oMat01.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value))
        //                            {
        //                                SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                                BubbleEvent = false;
        //                            }
        //                        }
        //                        if (pVal.ColUID == "ItemCode")
        //                        {
        //                            //UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                break;
        //        }

        //        //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "1")
        //                {
        //                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                    {
        //                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                        SubMain.Sbo_Application.ActivateMenuItem("1282");
        //                    }
        //                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                    {
        //                        PS_QM903_EnableFormItem();
        //                        PS_QM903_AddMatrixRow(1, oMat01.RowCount, ref false);
        //                        //oMat01
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                if (pVal.ItemChanged == true)
        //                {
        //                    //                    //헤더
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        PS_QM903_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //                    }

        //                    ////라인
        //                    if (pVal.ItemUID == "Mat01" & (pVal.ColUID == "ItmBsort" | pVal.ColUID == "ItmMsort" | pVal.ColUID == "CntcCode" | pVal.ColUID == "OrdNum" | pVal.ColUID == "ItemCode"))
        //                    {
        //                        PS_QM903_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                SubMain.RemoveForms(oFormUniqueID);
        //                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oForm = null;
        //                //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oMat01 = null;
        //                break;
        //        }
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgressBar01 = null;
        //    if (ErrNum == 101)
        //    {
        //        ErrNum = 0;
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //        BubbleEvent = false;
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;

        //    SAPbouiCOM.ComboBox oCombo = null;
        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1285":
        //                //복원
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }

        //        //-----------------------------------------------------------------------------------------------------------
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1285":
        //                //복원
        //                break;
        //            case "1293":
        //                //행삭제
        //                if (oMat01.RowCount != oMat01.VisualRowCount)
        //                {
        //                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //                    ////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
        //                    ////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
        //                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oMat01.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = i + 1;
        //                    }

        //                    oMat01.FlushToDataSource();
        //                    oDS_PS_QM903L.RemoveRecord(oDS_PS_QM903L.Size - 1);
        //                    //// Mat1에 마지막라인(빈라인) 삭제
        //                    oMat01.Clear();
        //                    oMat01.LoadFromDataSource();
        //                }
        //                break;

        //            case "1281":
        //                //찾기
        //                PS_QM903_EnableFormItem();
        //                oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                break;

        //            case "1282":
        //                //추가
        //                PS_QM903_EnableFormItem();
        //                PS_QM903_SetDocEntry();
        //                oDS_PS_QM903H.SetValue("U_DocDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD"));
        //                PS_QM903_AddMatrixRow(1, 0, ref true);
        //                //oMat01

        //                ////-- Combo Box 초기화
        //                //// 사업장
        //                oCombo = oForm.Items.Item("BPLId").Specific;
        //                oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

        //                //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oCombo = null;
        //                break;

        //            //                oForm.Items("DocDate").Click ct_Regular

        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                PS_QM903_EnableFormItem();
        //                if (oMat01.VisualRowCount > 0)
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(oMat01.VisualRowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01.VisualRowCount).Specific.Value))
        //                    {
        //                        PS_QM903_AddMatrixRow(1, oMat01.RowCount, ref false);
        //                    }
        //                }
        //                break;

        //        }
        //    }
        //    return;
        //Raise_MenuEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if ((eventInfo.BeforeAction == true))
        //    {

        //    }
        //    else if ((eventInfo.BeforeAction == false))
        //    {
        //        ////작업
        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((BusinessObjectInfo.BeforeAction == true))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((BusinessObjectInfo.BeforeAction == false))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //    }
        //    return;
        //Raise_FormDataEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion


    }
}
