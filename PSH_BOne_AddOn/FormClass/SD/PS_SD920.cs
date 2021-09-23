using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;
using System.Timers;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// A/R송장 미처리 납품 처리
	/// </summary>
	internal class PS_SD920 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_SD920H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD920L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        //DI API용 정보 클래스
        public class ItemInformation
        {
            public int LineNum;
            public string ODLNDocB;
            public string DLN1LinB;
            public string CardCode;
            public string CardName;
            public string ItemCode;
            public string ItemName;
            public string WhsCode;
            public double Quantity;
            public double OpenQty;
            public double Price;
            public double LineTotal;
            public double Qty;
            public string BaseType;
            public string BaseEntry;
            public string BaseLine;
            public string ULineNum;
            public int ORDRDoc;
            public int RDR1Line;
            public bool Check;
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD920.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD920_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD920");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_SD920_CreateItems();
                PS_SD920_SetComboBox();
                PS_SD920_SetInitial();
                PS_SD920_SetDocEntry();
                PS_SD920_EnableFormItem();

                oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
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
        private void PS_SD920_CreateItems()
        {
            try
            {
                oDS_PS_SD920H = oForm.DataSources.DBDataSources.Item("@PS_SD920H");
                oDS_PS_SD920L = oForm.DataSources.DBDataSources.Item("@PS_SD920L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD920_SetComboBox()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                oRecordSet01.DoQuery("SELECT BPLId, BPLName From[OBPL] order by 1");
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
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
        private void PS_SD920_SetInitial()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD920_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD920'", "");
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
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_SD920_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YYYYMM").Enabled = true;
                    oMat01.Columns.Item("Check").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YYYYMM").Enabled = true;
                    oMat01.Columns.Item("Check").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("YYYYMM").Enabled = false;
                    oMat01.Columns.Item("Check").Editable = false;
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
        private void PS_SD920_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_SD920L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD920L.Offset = oRow;
                oDS_PS_SD920L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD920_DeleteHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD920H.GetValue("U_YYYYMM", 0)))
                {
                    errMessage = "전기년월은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD920H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
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
        private bool PS_SD920_DeleteMatrixSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;            

            try
            {
                oMat01.FlushToDataSource();

                //라인
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
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
        /// 데이터 조회
        /// </summary>
        private void PS_SD920_LoadData()
        {
            string sQry;
            string YYYYMM;
            string BPLId;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_SD920_01] '" + BPLId + "','" + YYYYMM + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_SD920L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    throw new Exception();
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_SD920L.Size)
                    {
                        oDS_PS_SD920L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_SD920L.Offset = i;
                    oDS_PS_SD920L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD920L.SetValue("U_ODLNDocB", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_DLN1LinB", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_CardName", i, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("Dscription").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_WhsCode", i, oRecordSet01.Fields.Item("WhsCode").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_Quantity", i, oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_OpenQty", i, oRecordSet01.Fields.Item("OpenQty").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_Price", i, oRecordSet01.Fields.Item("Price").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_LinTotal", i, oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_Qty", i, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_BaseType", i, oRecordSet01.Fields.Item("U_BaseType").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_BaseDoc", i, oRecordSet01.Fields.Item("U_BaseEntry").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_BaseLine", i, oRecordSet01.Fields.Item("U_BaseLine").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_ULineNum", i, oRecordSet01.Fields.Item("U_LineNum").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_ORDRDoc", i, oRecordSet01.Fields.Item("BaseEntry").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_RDR1Line", i, oRecordSet01.Fields.Item("BaseLine").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

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
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 반품, 납품 DI API
        /// </summary>
        /// <returns></returns>
        private bool PS_SD920_AddoReturnsoDeliveryNotes()
        {
            bool returnValue = false;
            int i;
            int j;
            int K;
            int l = 0;
            int m;
            int LineCounter;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int RetVal;
            string sQry;
            //전체 StockInfo 구조체배열의 RowCount
            int[] LineNum = new int[1001];

            System.DateTime ORDNDocDate;
            System.DateTime ODLNDocDate;
            string ORDNDocEntry = null;
            string ODLNDocEntry = null;
            string BatchYN;
            SAPbobsCOM.Documents DI_oReturns = null; //반품 문서객체
            SAPbobsCOM.Documents DI_oDeliveryNotes = null; //납품 문서객체
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            Timer timer = new Timer();

            try
            {
                timer.Interval = 30000; //30초
                timer.Elapsed += KeepAddOnConnection;
                timer.Start();

                PSH_Globals.oCompany.StartTransaction();

                // 현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //반품,납품 대상

                oMat01.FlushToDataSource();

                ODLNDocDate = Convert.ToDateTime(codeHelpClass.Left(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 4) + '-' + codeHelpClass.Right(oDS_PS_SD920H.GetValue("U_YYYYMM", 0), 2) + "-01").AddMonths(1); //납품전기일(전기년월의 다음달 1일)
                ORDNDocDate = ODLNDocDate.AddDays(-1); //반품전기일(전기년월의 말일)

                for (i = 0; i <= oMat01.RowCount - 1; i++)
                {
                    if (oDS_PS_SD920L.GetValue("U_Check", i).ToString().Trim() == "Y")
                    {
                        ItemInformation itemInfo = new ItemInformation
                        {
                            LineNum = Convert.ToInt32(oDS_PS_SD920L.GetValue("U_LineNum", i).ToString().Trim()),
                            ODLNDocB = oDS_PS_SD920L.GetValue("U_ODLNDocB", i).ToString().Trim(),
                            DLN1LinB = oDS_PS_SD920L.GetValue("U_DLN1LinB", i).ToString().Trim(),
                            CardCode = oDS_PS_SD920L.GetValue("U_CardCode", i).ToString().Trim(),
                            CardName = oDS_PS_SD920L.GetValue("U_CardName", i).ToString().Trim(),
                            ItemCode = oDS_PS_SD920L.GetValue("U_ItemCode", i).ToString().Trim(),
                            ItemName = oDS_PS_SD920L.GetValue("U_ItemName", i).ToString().Trim(),
                            WhsCode = oDS_PS_SD920L.GetValue("U_WhsCode", i).ToString().Trim(),
                            Quantity = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_OpenQty", i).ToString().Trim()),
                            OpenQty = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_OpenQty", i).ToString().Trim()),
                            Price = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_Price", i).ToString().Trim()),
                            LineTotal = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_LinTotal", i).ToString().Trim()),
                            Qty = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_Qty", i).ToString().Trim()),
                            BaseType = oDS_PS_SD920L.GetValue("U_BaseType", i).ToString().Trim(),
                            BaseEntry = oDS_PS_SD920L.GetValue("U_BaseDoc", i).ToString().Trim(),
                            BaseLine = oDS_PS_SD920L.GetValue("U_BaseLine", i).ToString().Trim(),
                            ULineNum = oDS_PS_SD920L.GetValue("U_ULineNum", i).ToString().Trim(),
                            ORDRDoc = Convert.ToInt32(oDS_PS_SD920L.GetValue("U_ORDRDoc", i).ToString().Trim()),
                            RDR1Line = Convert.ToInt32(oDS_PS_SD920L.GetValue("U_RDR1Line", i).ToString().Trim()),
                            Check = false
                        };

                        itemInfoList.Add(itemInfo);
                    }
                }
                
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("반품 납품 생성!", itemInfoList.Count, false);
                
                string lclDocCur;
                double lclDocRate;
                string lclQuery;
                
                for (i = 0; i < itemInfoList.Count; i++)
                {
                    DI_oReturns = null;
                    DI_oDeliveryNotes = null;
                    DI_oReturns = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
                    DI_oDeliveryNotes = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);

                    K = i;
                    LineNum[i + 1] = itemInfoList[i].LineNum - 1;

                    sQry = "Select Manbtchnum From [OITM] Where ItemCode = '" + itemInfoList[i].ItemCode + "'";
                    oRecordSet01.DoQuery(sQry);
                    BatchYN = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                    if (BatchYN == "Y")
                    {
                        sQry = "  Select  BatchNum, ";
                        sQry += "         Quantity ";
                        sQry += " From    [IBT1] ";
                        sQry += " Where   BaseType = '15' ";
                        sQry += "         And BaseEntry = '" + itemInfoList[i].ODLNDocB + "'";
                        sQry += "         And BaseLinNum = '" + itemInfoList[i].DLN1LinB + "'";

                        oRecordSet01.DoQuery(sQry);
                    }

                    //반품_S
                    DI_oReturns.CardCode = itemInfoList[i].CardCode.Trim();
                    DI_oReturns.DocDate = ORDNDocDate;
                    DI_oReturns.DocDueDate = ORDNDocDate;
                    DI_oReturns.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());
                    DI_oReturns.Comments = "이월반품처리 [납품 : " + itemInfoList[i].ODLNDocB.Trim() + "]";

                    //헤더 환율처리_S (2011.11.04 송명규 추가)
                    lclQuery = "  SELECT  Currency,";
                    lclQuery += "         Rate";
                    lclQuery += " FROM    DLN1 ";
                    lclQuery += " WHERE   DocEntry = " + itemInfoList[i].ODLNDocB.Trim() + " ";
                    lclQuery += "         AND LineNum = " + itemInfoList[i].DLN1LinB.Trim();

                    oRecordSet02.DoQuery(lclQuery);

                    lclDocCur = oRecordSet02.Fields.Item("Currency").Value.ToString().Trim();
                    lclDocRate = Convert.ToDouble(oRecordSet02.Fields.Item("Rate").Value.ToString().Trim());

                    if (lclDocCur != "KRW")
                    {
                        DI_oReturns.DocCurrency = lclDocCur;
                        DI_oReturns.DocRate = lclDocRate;
                    }
                    //헤더 환율처리_E

                    DI_oReturns.Lines.ItemCode = itemInfoList[i].ItemCode;
                    DI_oReturns.Lines.Quantity = itemInfoList[i].OpenQty;
                    DI_oReturns.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    DI_oReturns.Lines.UnitPrice = itemInfoList[i].Price;
                    DI_oReturns.Lines.LineTotal = itemInfoList[i].LineTotal;
                    DI_oReturns.Lines.BaseType = 15;
                    DI_oReturns.Lines.BaseEntry = Convert.ToInt32(itemInfoList[i].ODLNDocB);
                    DI_oReturns.Lines.BaseLine = Convert.ToInt32(itemInfoList[i].DLN1LinB);
                    DI_oReturns.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                    DI_oReturns.Lines.UserFields.Fields.Item("U_BaseType").Value = itemInfoList[i].BaseType; //기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oReturns.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[i].BaseEntry; //기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oReturns.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[i].BaseLine; //기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                    //라인 환율처리_S
                    if (lclDocCur != "KRW")
                    {
                        DI_oReturns.Lines.Currency = lclDocCur;
                        DI_oReturns.Lines.Rate = lclDocRate;
                    }
                    //라인 환율처리_E

                    if (BatchYN == "Y")
                    {
                        m = 0;
                        while (!oRecordSet01.EoF)
                        {
                            if (m > 0)
                            {
                                DI_oReturns.Lines.BatchNumbers.Add();
                            }

                            DI_oReturns.Lines.BatchNumbers.BatchNumber = oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim();
                            DI_oReturns.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                            //oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                            //oRecordSet01.Fields("Quantity").Value 로 재 수정 20200511 황영수
                            oRecordSet01.MoveNext();
                            m += 1;
                        }
                    }
                    //반품_E

                    //납품_S
                    DI_oDeliveryNotes.CardCode = itemInfoList[i].CardCode.Trim();
                    DI_oDeliveryNotes.DocDate = ODLNDocDate;
                    DI_oDeliveryNotes.DocDueDate = ODLNDocDate;
                    DI_oDeliveryNotes.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());
                    DI_oDeliveryNotes.Comments = "이월반품처리 [판매오더 : " + itemInfoList[i].ODLNDocB.Trim() + "]";

                    //헤더 환율처리_S (2011.11.04 송명규 추가)
                    if (lclDocCur != "KRW")
                    {
                        DI_oDeliveryNotes.DocCurrency = lclDocCur;
                        DI_oDeliveryNotes.DocRate = lclDocRate;
                    }
                    //헤더 환율처리_E

                    DI_oDeliveryNotes.Lines.ItemCode = itemInfoList[i].ItemCode;
                    DI_oDeliveryNotes.Lines.Quantity = itemInfoList[i].OpenQty;
                    DI_oDeliveryNotes.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    DI_oDeliveryNotes.Lines.UnitPrice = itemInfoList[i].Price;
                    DI_oDeliveryNotes.Lines.LineTotal = itemInfoList[i].LineTotal;
                    //9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
                    //        If oS_PS_SD920L(i).ORDRDoc > 0 Then
                    //            DI_oDeliveryNotes.Lines.BaseType = 17
                    //            DI_oDeliveryNotes.Lines.BaseEntry = oS_PS_SD920L(i).ORDRDoc
                    //            DI_oDeliveryNotes.Lines.BaseLine = oS_PS_SD920L(i).RDR1Line
                    //        End If
                    //9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseType").Value = itemInfoList[i].BaseType; //기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[i].BaseEntry; //기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[i].BaseLine; //기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                    //라인 환율처리_S
                    if (lclDocCur != "KRW")
                    {
                        DI_oDeliveryNotes.Lines.Currency = lclDocCur;
                        DI_oDeliveryNotes.Lines.Rate = lclDocRate;
                    }
                    //라인 환율처리_E

                    if (BatchYN == "Y")
                    {
                        m = 0;
                        oRecordSet01.MoveFirst();
                        while (!oRecordSet01.EoF)
                        {
                            if (m > 0)
                            {
                                DI_oDeliveryNotes.Lines.BatchNumbers.Add();
                            }
                            DI_oDeliveryNotes.Lines.BatchNumbers.BatchNumber = oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim();
                            DI_oDeliveryNotes.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                            //oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                            //oRecordSet01.Fields("Quantity").Value 로 재 수정 20200511 황영수
                            //oRecordSet01.Fields("Quantity").Value 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").Value
                            oRecordSet01.MoveNext();
                            m += 1;
                        }
                    }
                    Add_Line_Again:
                    //납품_E

                    if (i != itemInfoList.Count - 1)
                    {
                        if (itemInfoList[i].ODLNDocB == itemInfoList[i + 1].ODLNDocB)
                        {
                            i += 1;
                            LineNum[i + 1] = itemInfoList[i].LineNum - 1;

                            sQry = "Select Manbtchnum From [OITM] Where ItemCode = '" + itemInfoList[i].ItemCode + "'";
                            oRecordSet01.DoQuery(sQry);
                            BatchYN = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            if (BatchYN == "Y")
                            {
                                sQry = "  Select    BatchNum,";
                                sQry += "           Quantity";
                                sQry += " From      [IBT1] ";
                                sQry += " Where     BaseType = '15'";
                                sQry += "           And BaseEntry = '" + itemInfoList[i].ODLNDocB.Trim() + "'";
                                sQry += "           And BaseLinNum = '" + itemInfoList[i].DLN1LinB.Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                            }

                            //반품
                            DI_oReturns.Lines.Add();
                            DI_oReturns.Lines.ItemCode = itemInfoList[i].ItemCode;
                            DI_oReturns.Lines.Quantity = itemInfoList[i].OpenQty;
                            DI_oReturns.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                            DI_oReturns.Lines.UnitPrice = itemInfoList[i].Price;
                            DI_oReturns.Lines.LineTotal = itemInfoList[i].LineTotal;
                            DI_oReturns.Lines.BaseType = 15;
                            DI_oReturns.Lines.BaseEntry = Convert.ToInt32(itemInfoList[i].ODLNDocB);
                            DI_oReturns.Lines.BaseLine = Convert.ToInt32(itemInfoList[i].DLN1LinB);
                            DI_oReturns.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                            DI_oReturns.Lines.UserFields.Fields.Item("U_BaseType").Value = itemInfoList[i].BaseType; //기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                            DI_oReturns.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[i].BaseEntry; //기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                            DI_oReturns.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[i].BaseLine; //기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                            //환율처리_S
                            if (lclDocCur != "KRW")
                            {
                                DI_oReturns.Lines.Currency = lclDocCur;
                                DI_oReturns.Lines.Rate = lclDocRate;
                            }
                            //환율처리_E

                            if (BatchYN == "Y")
                            {
                                m = 0;
                                while (!oRecordSet01.EoF)
                                {
                                    if (m > 0)
                                    {
                                        DI_oReturns.Lines.BatchNumbers.Add();
                                    }
                                    DI_oReturns.Lines.BatchNumbers.BatchNumber = oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim();
                                    DI_oReturns.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                                    //oRecordSet01.Fields("Quantity").Value 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                                    //oRecordSet01.Fields("Quantity").Value 로 재 수정 20200511 황영수
                                    //oRecordSet01.Fields("Quantity").Value 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").Value
                                    oRecordSet01.MoveNext();
                                    m += 1;
                                }
                            }

                            //납품
                            DI_oDeliveryNotes.Lines.Add();
                            DI_oDeliveryNotes.Lines.ItemCode = itemInfoList[i].ItemCode;
                            DI_oDeliveryNotes.Lines.Quantity = itemInfoList[i].OpenQty;
                            DI_oDeliveryNotes.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                            DI_oDeliveryNotes.Lines.UnitPrice = itemInfoList[i].Price;
                            DI_oDeliveryNotes.Lines.LineTotal = itemInfoList[i].LineTotal;
                            //9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
                            //                If oS_PS_SD920L(i).ORDRDoc > 0 Then
                            //                    DI_oDeliveryNotes.Lines.BaseType = 17
                            //                    DI_oDeliveryNotes.Lines.BaseEntry = oS_PS_SD920L(i).ORDRDoc
                            //                    DI_oDeliveryNotes.Lines.BaseLine = oS_PS_SD920L(i).RDR1Line
                            //                End If
                            //9.2 버전에서는 코드 오류 남(이미 마감된 판매오더 번호는 등록 불가한 것 같음, 따라서 주석 처리(2018.03.07 송명규))
                            DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                            DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseType").Value = itemInfoList[i].BaseType; //기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                            DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[i].BaseEntry; //기준납품문서의 납품처리[SD040]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                            DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[i].BaseLine; //기준납품문서의 납품처리[SD040]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                            //환율처리_S
                            if (lclDocCur != "KRW")
                            {
                                DI_oDeliveryNotes.Lines.Currency = lclDocCur;
                                DI_oDeliveryNotes.Lines.Rate = lclDocRate;
                            }
                            //환율처리_E

                            oRecordSet01.MoveFirst();
                            if (BatchYN == "Y")
                            {
                                m = 0;
                                while (!oRecordSet01.EoF)
                                {
                                    if (m > 0)
                                    {
                                        DI_oDeliveryNotes.Lines.BatchNumbers.Add();
                                    }
                                    DI_oDeliveryNotes.Lines.BatchNumbers.BatchNumber = oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim();
                                    DI_oDeliveryNotes.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                                    //oRecordSet01.Fields("Quantity").Value 'oS_PS_SD920L(i).OpenQty 가용재고로 입력되도록 수정 20170307
                                    //oRecordSet01.Fields("Quantity").Value 로 재 수정 20200511 황영수
                                    //oRecordSet01.Fields("Quantity").Value 'oS_PS_SD920L(i).OpenQty 'oRecordSet01.Fields("Quantity").Value
                                    oRecordSet01.MoveNext();
                                    m += 1;
                                }
                            }

                            goto Add_Line_Again;
                        }
                        else
                        {
                            goto Add_DI_oReturns_oDeliveryNotes;
                        }
                    }
                    else
                    {
                        goto Add_DI_oReturns_oDeliveryNotes;
                    }
                    Add_DI_oReturns_oDeliveryNotes:

                    //반품 완료
                    if (DI_oReturns != null)
                    {
                        RetVal = DI_oReturns.Add();
                        if (0 != RetVal)
                        {
                            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                            errCode = "1";
                            throw new Exception();
                        }

                        PSH_Globals.oCompany.GetNewObjectCode(out ORDNDocEntry);

                        sQry = "EXECUTE [PS_Z_RETU_GI_ZSD920] '" + ORDNDocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }

                    //납품 완료
                    if (DI_oDeliveryNotes != null)
                    {
                        RetVal = DI_oDeliveryNotes.Add();
                        if (0 != RetVal)
                        {
                            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                            errCode = "1";
                            throw new Exception();
                        }

                        PSH_Globals.oCompany.GetNewObjectCode(out ODLNDocEntry);
                    }

                    LineCounter = 0;
                    for (j = K; j <= i; j++)
                    {
                        oDS_PS_SD920L.SetValue("U_ORDNDoc", LineNum[l + 1], ORDNDocEntry);
                        oDS_PS_SD920L.SetValue("U_RDN1Line", LineNum[l + 1], Convert.ToString(LineCounter));
                        oDS_PS_SD920L.SetValue("U_ODLNDoc", LineNum[l + 1], ODLNDocEntry);
                        oDS_PS_SD920L.SetValue("U_DLN1Line", LineNum[l + 1], Convert.ToString(LineCounter));
                        //납품문서 DLN1 (BaseEntry, BaseLine,BaseType 추가함)
                        sQry = "EXECUTE [PS_Z_InsertInfo_ODLN] '" + ODLNDocEntry + "','" + LineCounter + "','" + itemInfoList[l].ORDRDoc.ToString().Trim() + "','" + itemInfoList[l].RDR1Line.ToString().Trim() + "'";

                        oRecordSet01.DoQuery(sQry);

                        LineCounter += 1;
                        l += 1;
                    }

                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + itemInfoList.Count + "건의 납품 반품 문서 생성중...!";
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
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
                timer.Stop();
                timer.Dispose();

                if (DI_oReturns != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oReturns);
                }

                if (DI_oDeliveryNotes != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oDeliveryNotes);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            }

            return returnValue;
        }

        /// <summary>
        /// AddOn 연결 유지용 Timer 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void KeepAddOnConnection(object sender, ElapsedEventArgs e)
        {
            PSH_Globals.SBO_Application.RemoveWindowsMessage(BoWindowsMessageType.bo_WM_TIMER, true);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD920_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_SD920_DeleteMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (PS_SD920_AddoReturnsoDeliveryNotes() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_SD920_DeleteHeaderSpaceLine() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PS_SD920_LoadData();
                    }
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
                    if (pVal.ItemUID == "Mat01" && pVal.Row != 0 && pVal.ColUID == "Check")
                    {
                        oForm.Freeze(true);
                        oMat01.FlushToDataSource();
                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PS_SD920L.GetValue("U_ODLNDocB", i) == oDS_PS_SD920L.GetValue("U_ODLNDocB", pVal.Row - 1))
                            {
                                oDS_PS_SD920L.SetValue("U_Check", i, oDS_PS_SD920L.GetValue("U_Check", pVal.Row - 1).ToString().Trim());
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oForm.Freeze(false);
                    }
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
            string check = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "Check")
                    {
                        oForm.Freeze(true);
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_SD920L.GetValue("U_Check", 0).ToString().Trim()) || oDS_PS_SD920L.GetValue("U_Check", 0).ToString().Trim() == "N")
                        {
                            check = "Y";
                        }
                        else if (oDS_PS_SD920L.GetValue("U_Check", 0).ToString().Trim() == "Y")
                        {
                            check = "N";
                        }
                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            oDS_PS_SD920L.SetValue("U_Check", i, check);
                        }
                        oMat01.LoadFromDataSource();
                        oForm.Freeze(false);
                    }
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
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_SD920_AddMatrixRow(oMat01.RowCount, false);
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD920H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD920L);
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
                            oForm.Freeze(true);
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_SD920L.RemoveRecord(oDS_PS_SD920L.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PS_SD920_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            oForm.Freeze(false);
                            break;
                        case "1281": //찾기
                            oForm.Freeze(true);
                            PS_SD920_EnableFormItem();
                            oForm.Freeze(false);
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_SD920_EnableFormItem();
                            PS_SD920_SetInitial();
                            PS_SD920_SetDocEntry();
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            oForm.Freeze(true);
                            PS_SD920_EnableFormItem();
                            oForm.Freeze(false);
                            break;
                        case "1287": //복제
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
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
            }
        }
    }
}