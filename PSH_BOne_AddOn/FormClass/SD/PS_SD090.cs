using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이동처리등록
	/// </summary>
	internal class PS_SD090 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_SD090H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD090L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oSeq;

		// 입고 DI를 위한 정보를 가지는 구조체
		public class StockInfos
		{
			public string CardCode; //고객코드
			public string ItemCode; //품목코드
			public string FromWarehouseCode; //창고코드
			public string ToWarehouseCode; //창고코드
			public double Weight; //중량
			public double UnWeight;
			public string BatchNum; //배치번호
			public double BatchWeight;//배치중량
			public int Qty; //수량
			public string TransNo; //재고이전문서번호
			public string Chk;
			public int MatrixRow;
			public string StockTransDocEntry; //재고이전문서번호
			public string StockTransLineNum; //재고이전라인번호
			public string Indate; //전기일
        }

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD090.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD090_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD090");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD090_CreateItems();
				PS_SD090_SetComboBox();
				PS_SD090_EnableFormItem();
				PS_SD090_SetDocEntry();
                PS_SD090_AddMatrixRow(0, true);

                oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1283", false); //제거
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1285", false); //복원
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
        private void PS_SD090_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_SD090H = oForm.DataSources.DBDataSources.Item("@PS_SD090H");
                oDS_PS_SD090L = oForm.DataSources.DBDataSources.Item("@PS_SD090L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();
                
                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                oDS_PS_SD090H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL Where BPLId = '1' Or BPLId = '4' order by BPLId", dataHelpClass.User_BPLID(), false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD090_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("OutWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "104", false, true); //출고창고
                dataHelpClass.Set_ComboList(oForm.Items.Item("InWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "101", false, true); //입고창고

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", ""); //품목대분류
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", ""); //품목중분류
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", ""); //형태타입
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code", "", ""); //질별
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", ""); //인증기호
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code", "", ""); //매입기준단위
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_SD090_EnableFormItem()
        {
            string lQuery;
            SAPbobsCOM.Recordset lRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("RepName").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ShipTo").Enabled = true;
                    oForm.Items.Item("CarCo").Enabled = true;
                    oForm.Items.Item("CarNo").Enabled = true;
                    oForm.Items.Item("ArrSite").Enabled = true;
                    oForm.Items.Item("Fare").Enabled = true;
                    oForm.Items.Item("Specific").Enabled = true;
                    oForm.Items.Item("ChulPrin").Enabled = false;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemGu").Editable = true;
                    oForm.Items.Item("DocDate").Specific.String = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("RepName").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ShipTo").Enabled = true;
                    oForm.Items.Item("CarCo").Enabled = true;
                    oForm.Items.Item("CarNo").Enabled = true;
                    oForm.Items.Item("ArrSite").Enabled = true;
                    oForm.Items.Item("Fare").Enabled = false;
                    oForm.Items.Item("Specific").Enabled = false;
                    oForm.Items.Item("ChulPrin").Enabled = true;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGu").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("RepName").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("ShipTo").Enabled = false;
                    oForm.Items.Item("CarCo").Enabled = false;
                    oForm.Items.Item("CarNo").Enabled = false;
                    oForm.Items.Item("ArrSite").Enabled = false;
                    oForm.Items.Item("Fare").Enabled = false;
                    oForm.Items.Item("Specific").Enabled = false;
                    oForm.Items.Item("ChulPrin").Enabled = true;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGu").Editable = false;

                    lQuery = "SELECT Status,Canceled FROM [@PS_SD090H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    lRecordSet.DoQuery(lQuery);
                    if (lRecordSet.Fields.Item(0).Value == "O")
                    {
                        oForm.Items.Item("Status").Specific.Value = "미결";
                    }
                    else if (lRecordSet.Fields.Item(0).Value == "C")
                    {
                        if (lRecordSet.Fields.Item(1).Value == "Y")
                        {
                            oForm.Items.Item("Status").Specific.Value = "취소";
                        }
                        else
                        {
                            oForm.Items.Item("Status").Specific.Value = "종료";
                        }
                    }
                }
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lRecordSet);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD090_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD090'", "");
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

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD090_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_SD090L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD090L.Offset = oRow;
                oDS_PS_SD090L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        private void PS_SD090_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                        {
                            oMat01.FlushToDataSource();
                            PS_SD090_AddMatrixRow(oMat01.RowCount, false);
                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_DelHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_DocDate", 0))) 
                {
                    errMessage = "거래일자는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_OutWhCd", 0)))
                {
                    errMessage = "출고창고는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_InWhCd", 0)))
                {
                    errMessage = "입고창고는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }

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
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_DelMatrixSpaceLine()
        {
            bool returnValue = false;
            int i = 0;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_SD090L.Offset = i;
                        if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_ItemCode", i)))
                        {
                            errMessage = "아이템 데이터는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_Qty", i)))
                        {
                            errMessage = "수량은 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if(string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i)))
                        {
                            errMessage = "이동요청 문서는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                    }

                    if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", oMat01.VisualRowCount - 1)))
                    {
                        oDS_PS_SD090L.RemoveRecord(oMat01.VisualRowCount - 1);
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
        /// 재고이전
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_TransStock()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int RetVal;
            int StockTransLineCounter;
            int Q;
            int i;
            int K;
            int r;
            int DocCnt = 0;
            string Chk1_Val;
            string sCur_ItemCode;
            string sNxt_ItemCode;
            string sCur_TrCardCode;
            string sCur_TrOutWhs;
            string sNxt_TrOutWhs;
            string sCur_TrInWhs;
            string sNxt_TrInWhs;

            SAPbobsCOM.StockTransfer oStockTrans = null;
            SAPbouiCOM.ProgressBar oPrgBar = null;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oPrgBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<StockInfos> stockInfoList = new List<StockInfos>(); //재고

                for (i = 0; i <= oMat01.RowCount - 1; i++)
                {
                    StockInfos stockInfo = new StockInfos
                    {
                        CardCode = oDS_PS_SD090H.GetValue("U_CardCode", 0).ToString().Trim(),
                        ItemCode = oDS_PS_SD090L.GetValue("U_ItemCode", i).ToString().Trim(),
                        FromWarehouseCode = oDS_PS_SD090H.GetValue("U_OutWhCd", 0).ToString().Trim(),
                        ToWarehouseCode = oDS_PS_SD090H.GetValue("U_InWhCd", 0).ToString().Trim(),
                        BatchNum = oDS_PS_SD090L.GetValue("U_BatchNum", i).ToString().Trim(),
                        Weight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Weight", i).ToString().Trim()), 2),
                        UnWeight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Unweight", i).ToString().Trim()), 2),
                        BatchWeight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Qty", i).ToString().Trim()), 2),
                        Qty = Convert.ToInt32(oDS_PS_SD090L.GetValue("U_Qty", i).ToString().Trim()),
                        TransNo = oForm.Items.Item("DocEntry").Specific.Value + (i + 1),
                        Chk = "N",
                        MatrixRow = i + 1,
                        Indate = oForm.Items.Item("DocDate").Specific.Value
                    };

                    stockInfoList.Add(stockInfo);
                }

                for (i = 0; i < stockInfoList.Count; i++)
                {
                    stockInfoList[i].StockTransDocEntry = "";
                }

                for (i = 0; i < stockInfoList.Count; i++)
                {
                    Chk1_Val = stockInfoList[i].Chk;

                    if (Chk1_Val != "N")
                    {
                        continue;
                    }

                    sCur_TrOutWhs = stockInfoList[i].FromWarehouseCode;

                    oStockTrans = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    oStockTrans.CardCode = stockInfoList[i].CardCode;
                    oStockTrans.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(stockInfoList[i].Indate, "-"));
                    oStockTrans.FromWarehouse = sCur_TrOutWhs;
                    oStockTrans.Comments = "재고이전" + oForm.Items.Item("DocEntry").Specific.Value + ".";

                    StockTransLineCounter = -1;

                    for (K = i; K < stockInfoList.Count; K++)
                    {
                        Chk1_Val = stockInfoList[K].Chk;

                        if (Chk1_Val != "N")
                        {
                            continue;
                        }
                        sCur_TrCardCode = stockInfoList[K].CardCode;
                        sNxt_TrOutWhs = stockInfoList[K].FromWarehouseCode;
                        sCur_ItemCode = stockInfoList[K].ItemCode;
                        sCur_TrInWhs = stockInfoList[K].ToWarehouseCode;

                        if (sCur_TrOutWhs != sNxt_TrOutWhs)
                        {
                            continue;
                        }

                        if (i != K)
                        {
                            oStockTrans.Lines.Add();
                        }
                        StockTransLineCounter += 1;

                        oStockTrans.Lines.ItemCode = stockInfoList[K].ItemCode;
                        oStockTrans.Lines.UserFields.Fields.Item("U_Qty").Value = stockInfoList[K].Qty;
                        oStockTrans.Lines.UserFields.Fields.Item("U_UnWeight").Value = stockInfoList[K].UnWeight;
                        oStockTrans.Lines.Quantity = stockInfoList[K].Qty;
                        oStockTrans.Lines.WarehouseCode = stockInfoList[K].ToWarehouseCode;
                        oStockTrans.Lines.UserFields.Fields.Item("U_BatchNum").Value = stockInfoList[K].BatchNum;
                        oStockTrans.Lines.BatchNumbers.BatchNumber = stockInfoList[K].BatchNum;
                        oStockTrans.Lines.BatchNumbers.Quantity = System.Math.Round(stockInfoList[K].BatchWeight, 2);
                        oStockTrans.Lines.BatchNumbers.Notes = "재고이전(Addon)";
                        stockInfoList[K].Chk = "Y";
                        stockInfoList[K].StockTransDocEntry = "Checked";
                        stockInfoList[K].StockTransLineNum = Convert.ToString(StockTransLineCounter);

                        for (Q = K + 1; Q < stockInfoList.Count; Q++)
                        {
                            Chk1_Val = stockInfoList[Q].Chk;

                            if (Chk1_Val != "N")
                            {
                                continue;
                            }

                            sNxt_TrOutWhs = stockInfoList[Q].FromWarehouseCode;
                            sNxt_ItemCode = stockInfoList[Q].ItemCode;
                            sNxt_TrInWhs = stockInfoList[Q].ToWarehouseCode;

                            if (sNxt_TrOutWhs == sCur_TrOutWhs & sCur_ItemCode == sNxt_ItemCode & sCur_TrInWhs == sNxt_TrInWhs)
                            {
                                if (dataHelpClass.GetValue("SELECT ManBatchNum FROM OITM WHERE ITEMCODE = ''", 0, 1) == "Y")
                                {
                                    oStockTrans.Lines.BatchNumbers.Add();
                                    oStockTrans.Lines.BatchNumbers.BatchNumber = stockInfoList[Q].BatchNum;
                                    oStockTrans.Lines.BatchNumbers.Quantity = System.Math.Round(stockInfoList[Q].BatchWeight, 2);
                                    oStockTrans.Lines.UserFields.Fields.Item("Quantity").Value = Convert.ToInt32(oStockTrans.Lines.UserFields.Fields.Item("Quantity").Value) + stockInfoList[Q].Qty;
                                    oStockTrans.Lines.Quantity = oStockTrans.Lines.Quantity + System.Math.Round(stockInfoList[Q].Weight, 2);
                                    oStockTrans.Lines.BatchNumbers.Notes = "재고이전(Addon)";
                                    stockInfoList[Q].Chk = "Y";
                                    stockInfoList[Q].StockTransDocEntry = "Checked";
                                    stockInfoList[Q].StockTransLineNum = Convert.ToString(StockTransLineCounter);
                                }
                            }
                        }
                    }

                    RetVal = oStockTrans.Add();

                    if (RetVal == 0)
                    {
                        DocCnt += 1;
                        PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                        for (r = 0; r < stockInfoList.Count; r++)
                        {
                            if (stockInfoList[r].StockTransDocEntry == "Checked")
                            {
                                stockInfoList[r].StockTransDocEntry = afterDIDocNum;
                                oDS_PS_SD090H.SetValue("U_StoTrDoc", 0, stockInfoList[r].StockTransDocEntry);
                            }
                        }
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errCode = "1";
                        throw new Exception();
                    }
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                PSH_Globals.SBO_Application.StatusBar.SetText(DocCnt + " 개의 재고이전 문서가 발행되었습니다 !", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oStockTrans != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTrans);
                }

                if (oPrgBar != null)
                {
                    oPrgBar.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPrgBar);
                }
            }

            return returnValue;
        }



        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string sQry = null;
        //	string sQry02 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	SAPbobsCOM.Recordset oRecordset02 = null;
        //	object TempForm01 = null;
        //	short ErrNum = 0;

        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//// 객체 정의 및 데이터 할당
        //	oRecordset02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


        //	short i = 0;
        //	short j = 0;
        //	short DocEntry = 0;
        //	short LineId = 0;
        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.EventType) {

        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pVal.ItemUID == "1") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (PS_SD090_DelHeaderSpaceLine() == false) {
        //							BubbleEvent = false;
        //							//BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
        //							return;
        //						}

        //						if (PS_SD090_DelMatrixSpaceLine() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						//// 재고 이동 DI API
        //						if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //							if (PS_SD090_TransStock() == true) {
        //								PS_SD090_UpdateUserField();
        //							} else {
        //								PS_SD090_AddMatrixRow(1, oMat01.RowCount, ref true);
        //								BubbleEvent = false;
        //								return;
        //							}
        //						}
        //					}

        //				} else if (pVal.ItemUID == "ChulPrin") {
        //					PS_SD090_Print_Report01();
        //					//            ElseIf pVal.ItemUID = "GuraPrin" Then

        //				} else {
        //					if (pVal.ItemChanged == true) {
        //						if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
        //							PS_SD090_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //						}
        //					}
        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2

        //				// 거래처코드
        //				//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value)) {
        //					if (pVal.ItemUID == "CardCode" & pVal.CharPressed == 9) {
        //						////CharPressed: The character that was pressed to trigger this event.
        //						oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}

        //				// 아이템코드
        //				if (pVal.ItemUID == "Mat01" & pVal.ColUID == "SD091HNo" & pVal.CharPressed == 9) {
        //					//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)) {
        //						oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}

        //				// 담당자
        //				//UPGRADE_WARNING: oForm.Items(RepName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item("RepName").Specific.Value)) {
        //					if (pVal.ItemUID == "RepName" & pVal.CharPressed == 9) {
        //						oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}

        //				// 납품처
        //				//UPGRADE_WARNING: oForm.Items(ShipTo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item("ShipTo").Specific.Value)) {
        //					if (pVal.ItemUID == "ShipTo" & pVal.CharPressed == 9) {
        //						oForm.Items.Item("ShipTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}

        //				// 운송업체

        //				// 차량번호

        //				// 도착장소

        //				// 질별

        //				// 출고창고
        //				if (pVal.ItemUID == "OutWhCd" & pVal.CharPressed == 9) {
        //					//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Specific.String)) {
        //						//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}

        //				// 입고창고
        //				if (pVal.ItemUID == "InWhCd" & pVal.CharPressed == 9) {
        //					//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Specific.String)) {
        //						//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
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
        //				if (pVal.ItemChanged == true) {
        //					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
        //						PS_SD090_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //					}
        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				if (pVal.ItemChanged == true) {

        //					// 거래처 이름 Query
        //					if (pVal.ItemUID == "CardCode") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "Select CardName From [OCRD] Where CardCode = '" + Strings.Trim(oForm.Items.Item("CardCode").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(CardName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("CardName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					}

        //					// 사원 이름 Query
        //					if (pVal.ItemUID == "RepName") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + Strings.Trim(oForm.Items.Item("RepName").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(RepNm1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("RepNm1").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					}

        //					// 납품처 이름 Query
        //					if (pVal.ItemUID == "ShipTo") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "Select CardName From [OCRD] Where CardCode = '" + Strings.Trim(oForm.Items.Item("ShipTo").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(ShipNm).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("ShipNm").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					}

        //					// 아이템 코드
        //					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
        //						PS_SD090_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //					}

        //					// 출고 창고
        //					if (pVal.ItemUID == "OutWhCd") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "Select WhsName From [OWHS] Where WhsCode = '" + Strings.Trim(oForm.Items.Item("OutWhCd").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(OutWhName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("OutWhName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					}

        //					// 입고 창고
        //					if (pVal.ItemUID == "InWhCd") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "Select WhsName From [OWHS] Where WhsCode = '" + Strings.Trim(oForm.Items.Item("InWhCd").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(InWhCd).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("InWhCd").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					}

        //					// 작업요청
        //					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "SD091HNo") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						j = Strings.InStr(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), "-");
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						DocEntry = Convert.ToInt16(Strings.Left(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), j - 1));
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						LineId = Convert.ToInt16(Strings.Mid(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), j + 1));
        //						sQry = "Select U_ItemCode, U_ItemName, U_ItemGu, U_Qty, ";
        //						sQry = sQry + "U_Unweight, U_Weight, U_Comments, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, ";
        //						sQry = sQry + "U_Quality, U_Mark, U_CallSize, U_SbasUnit ";
        //						sQry = sQry + "From [@PS_SD091L] Where DocEntry = '" + DocEntry + "' And LineId = '" + LineId + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						oMat01.FlushToDataSource();
        //						oDS_PS_SD090L.SetValue("U_ItemCode", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
        //						oDS_PS_SD090L.SetValue("U_ItemName", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(1).Value));
        //						oDS_PS_SD090L.SetValue("U_ItemGu", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(2).Value));
        //						oDS_PS_SD090L.SetValue("U_Qty", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(3).Value));
        //						oDS_PS_SD090L.SetValue("U_Unweight", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(4).Value));
        //						oDS_PS_SD090L.SetValue("U_Weight", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(5).Value));
        //						oDS_PS_SD090L.SetValue("U_Comments", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(6).Value));
        //						oDS_PS_SD090L.SetValue("U_ItmBsort", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(7).Value));
        //						oDS_PS_SD090L.SetValue("U_ItmMsort", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(8).Value));
        //						oDS_PS_SD090L.SetValue("U_Unit1", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(9).Value));
        //						oDS_PS_SD090L.SetValue("U_Size", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(10).Value));
        //						oDS_PS_SD090L.SetValue("U_ItemType", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(11).Value));
        //						oDS_PS_SD090L.SetValue("U_Quality", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(12).Value));
        //						oDS_PS_SD090L.SetValue("U_Mark", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(13).Value));
        //						oDS_PS_SD090L.SetValue("U_CallSize", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(14).Value));
        //						oDS_PS_SD090L.SetValue("U_SbasUnit", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(15).Value));
        //						//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oRecordSet01 = null;

        //						PS_SD090_AddMatrixRow(1, oMat01.VisualRowCount, ref false);
        //						for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //							if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i - 1))) {
        //								oDS_PS_SD090L.RemoveRecord(i - 1);
        //								oMat01.LoadFromDataSource();
        //							}
        //						}

        //						oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //						sQry = "Select U_OutWhCd, U_InWhCd From [@PS_SD091H] Where DocEntry = '" + DocEntry + "'";
        //						oRecordSet01.DoQuery(sQry);
        //						oDS_PS_SD090H.SetValue("U_OutWhCd", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
        //						oDS_PS_SD090H.SetValue("U_InWhCd", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(1).Value));
        //						//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oRecordSet01 = null;


        //						for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //							//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //								SumQty = SumQty;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //							}
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //						}
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SumWeight").Specific.Value = SumWeight;


        //						PS_SD090_AddMatrixRow(1, oMat01.VisualRowCount, ref false);
        //					}

        //					//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet01 = null;
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			//                PS_SD090_AddMatrixRow 1, oMat01.VisualRowCount, False
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
        //				oLast_Item_UID = pVal.ItemUID;
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				//저장 후 추가 가능처리
        //				if (pVal.ItemUID == "1") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true) {
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						SubMain.Sbo_Application.ActivateMenuItem("1282");
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == false) {
        //						PS_SD090_EnableFormItem();
        //						PS_SD090_AddMatrixRow(1, oMat01.RowCount, ref true);
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				if (pVal.Action_Success == true) {
        //					oSeq = 1;
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				if (pVal.ItemChanged == true) {
        //					oForm.Freeze(true);
        //					if ((pVal.ItemUID == "BPLId")) {
        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (oForm.Items.Item("BPLId").Specific.Value == "1") {
        //							oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "101");
        //							oDS_PS_SD090H.SetValue("U_InWhCd", 0, "104");
        //							//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (oForm.Items.Item("BPLId").Specific.Value == "4") {
        //							oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "104");
        //							oDS_PS_SD090H.SetValue("U_InWhCd", 0, "101");
        //						}
        //					}
        //					oForm.Update();
        //					oForm.Freeze(false);
        //				}
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
        //			// 이동요청 문서 Query


        //			//            ' 품목 이름 Query
        //			//                If (pVal.ItemUID = "Mat01" And (pVal.ColUID = "ItemCode")) Then
        //			//                    sQry = "Select ItemName, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit From [OITM] Where "
        //			//                    sQry = sQry & "ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pVal.Row).Specific.Value) & "'"
        //			//                    oRecordSet01.DoQuery sQry
        //			//                    oMat01.Columns("ItemName").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        //			//            ' 품목 대분류
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(1).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//            ' 품목 중분류
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(2).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//            ' 형태타입
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordSet01.Fields(5).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("ItemType").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//            ' 질별
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordSet01.Fields(6).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("Quality").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//            ' 인증기호
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordSet01.Fields(7).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("Mark").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//            ' 판매기준단위
        //			//                    sQry02 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordSet01.Fields(9).Value) & "'"
        //			//                    oRecordSet02.DoQuery sQry02
        //			//                    Call oMat01.Columns("SbasUnit").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
        //			//                    oMat01.Columns("Unit1").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(3).Value)
        //			//                    oMat01.Columns("Size").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(4).Value)
        //			//                End If
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11

        //				for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //						SumQty = SumQty;
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //					}
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //				}

        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

        //				PS_SD090_AddMatrixRow(1, oMat01.VisualRowCount, ref true);
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				if (oSeq == 1) {
        //					oSeq = 0;
        //				}
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
        //				oLast_Item_UID = pVal.ItemUID;
        //				break;
        //			//                If oLast_Item_UID = "매트릭스" Then
        //			//                    If pVal.Row > 0 Then
        //			//                        oLast_Item_UID = pVal.ItemUID
        //			//                        oLast_Col_UID = pVal.ColUID
        //			//                        oLast_Col_Row = pVal.Row
        //			//                    End If
        //			//                Else
        //			//                    oLast_Item_UID = pVal.ItemUID
        //			//                    oLast_Col_UID = ""
        //			//                    oLast_Col_Row = 0
        //			//                End If
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				break;
        //		}
        //	}

        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행닫기
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
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
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_SD090_EnableFormItem();
        //				break;
        //			//                oForm.Items("ItemCode").Click ct_Regular
        //			case "1282":
        //				//추가
        //				PS_SD090_EnableFormItem();
        //				PS_SD090_SetDocEntry();
        //				PS_SD090_AddMatrixRow(0, 0, ref true);
        //				oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //				break;

        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_SD090_EnableFormItem();
        //				if (oMat01.VisualRowCount > 0) {
        //					//UPGRADE_WARNING: oMat01.Columns(SD091HNo).Cells(oMat01.VisualRowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (!string.IsNullOrEmpty(oMat01.Columns.Item("SD091HNo").Cells.Item(oMat01.VisualRowCount).Specific.Value)) {
        //						PS_SD090_AddMatrixRow(1, oMat01.RowCount, ref true);
        //					}
        //				}
        //				break;
        //			case "1293":
        //				//행닫기
        //				if (oMat01.RowCount != oMat01.VisualRowCount) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					// DBDataSource에 레코드가 한줄 더 생긴다.
        //					oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1);
        //					// 레코드 한 줄을 지운다.
        //					oMat01.LoadFromDataSource();
        //					// DBDataSource를 매트릭스에 올리고
        //					if (oMat01.RowCount == 0) {
        //						PS_SD090_AddMatrixRow(1, 0, ref true);
        //					} else {
        //						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
        //							PS_SD090_AddMatrixRow(1, oMat01.RowCount, ref true);

        //						}
        //					}


        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //							SumQty = SumQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //				}
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34 - 추가
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35 - 업데이트
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
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
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((eventInfo.BeforeAction == true)) {
        //		////작업
        //	} else if ((eventInfo.BeforeAction == false)) {
        //		////작업
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if (pVal.BeforeAction == true) {
        //		if (SubMain.Sbo_Application.MessageBox("정말 삭제 하시겠습니까?", 1, "OK", "NO") != 1) {
        //			BubbleEvent = false;
        //		}
        //		////행삭제전 행삭제가능여부검사
        //	} else if (pVal.BeforeAction == false) {
        //		for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //		}
        //		//        oMat01.Clear
        //		oMat01.FlushToDataSource();
        //		oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1);
        //		oMat01.LoadFromDataSource();
        //		if (oMat01.RowCount == 0) {
        //			PS_SD090_AddMatrixRow(0, oMat01.RowCount, ref true);
        //		} else {
        //			if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
        //				PS_SD090_AddMatrixRow(1, oMat01.RowCount, ref true);
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion



        #region PS_SD090_Print_Report01
        //private void PS_SD090_Print_Report01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry01 = null;

        //	MDC_PS_Common.ConnectODBC();
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	WinTitle = "[PS_SD090] 출고원부/반출증";
        //	ReportName = "PS_SD090_10.rpt";
        //	sQry01 = "EXEC PS_SD090_10 '" + DocEntry + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry01, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}
        //	return;
        //	PS_SD090_Print_Report01_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD090_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}
