using System;
using System.Linq;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;
using System.Timers;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 이월반품처리(원가) 
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
        private string DocEntry;

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
            public string BatchNum;
            public double BatchQty;
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

        public class ItemInformation2
        {
            public string ODLNDocB;
            public string DLN1LinB;
            public string ORDNDocEnt;
            public string RDN1Lin;
            public string ODLNDocEnt;
            public string DLN1Lin;
        }

        List<ItemInformation2> itemInfoList2 = new List<ItemInformation2>();

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
                oRecordSet01.DoQuery("SELECT BPLId, BPLName From[OBPL] order by 1"); //사업장
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
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YYYYMM").Enabled = true;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("YYYYMM").Enabled = false;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가    
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD920H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD920H.GetValue("U_YYYYMM", 0)))
                {
                    errMessage = "전기년월은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                // 마감일자 Check
                else if (dataHelpClass.Check_Finish_Status(oDS_PS_SD920H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_SD920H.GetValue("U_YYYYMM", 0).ToString().Trim()) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기년월을 확인하고, 회계부서로 문의하세요.";
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
                if (oMat01.VisualRowCount == 0) //라인
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
                    oDS_PS_SD920L.SetValue("U_BatchNum", i, oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim());
                    oDS_PS_SD920L.SetValue("U_BatchQty", i, oRecordSet01.Fields.Item("BatchQty").Value.ToString().Trim());
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
            int j = 0;
            int k = 0;
            int l = 0; // itemInfoList2 행            
            int errDICode = 0;
            int RetVal;
            int RowCnt = 0;
            int subRowCnt = 0; //납품문서 행     
            int lineCnt = 0; //반품,납품 문서라인 번호  
            double lclDocRate;
            string lclDocCur;
            string lclQuery;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            string sQry;
            System.DateTime ORDNDocDate;
            System.DateTime ODLNDocDate;
            string ORDNDocEntry = null;
            string ODLNDocEntry = null;            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Documents DI_oReturns = null;
            SAPbobsCOM.Documents DI_oDeliveryNotes = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            Timer timer = new Timer();

            try
            {
                timer.Interval = 30000; //30초
                timer.Elapsed += KeepAddOnConnection;
                timer.Start();
                oMat01.FlushToDataSource();

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                // 현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //반품,납품 대상

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
                            BatchNum = oDS_PS_SD920L.GetValue("U_BatchNum", i).ToString().Trim(),
                            BatchQty = Convert.ToDouble(oDS_PS_SD920L.GetValue("U_BatchQty", i).ToString().Trim()),
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
                
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("반품 납품 생성중!", itemInfoList.Count, false);

                for (i = 0; i < itemInfoList.Count; i++, j++)
                {
                    if (DI_oReturns == null && DI_oDeliveryNotes == null)
                    {
                        DI_oReturns = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns); //반품 문서객체
                        DI_oDeliveryNotes = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes); //납품 문서객체 
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
                    DI_oReturns.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    DI_oReturns.Lines.UnitPrice = itemInfoList[i].Price;
                    DI_oReturns.Lines.BaseType = 15;
                    DI_oReturns.Lines.BaseEntry = Convert.ToInt32(itemInfoList[i].ODLNDocB);
                    DI_oReturns.Lines.BaseLine = Convert.ToInt32(itemInfoList[i].DLN1LinB);
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
                    if (!string.IsNullOrEmpty(itemInfoList[i].BatchNum))
                    {
                        DI_oReturns.Lines.Quantity = itemInfoList[i].OpenQty;
                        var linQry = from c in itemInfoList
                                     where c.ODLNDocB == itemInfoList[i].ODLNDocB && c.ItemCode == itemInfoList[i].ItemCode && c.DLN1LinB == itemInfoList[i].DLN1LinB
                                     select c.ODLNDocB.Count();
                        RowCnt = linQry.Count();
                        for (int BatchCnt = i; BatchCnt < i + RowCnt; BatchCnt++)
                        {
                            DI_oReturns.Lines.BatchNumbers.BatchNumber = itemInfoList[BatchCnt].BatchNum;
                            DI_oReturns.Lines.BatchNumbers.Quantity = itemInfoList[BatchCnt].BatchQty;
                            DI_oReturns.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[BatchCnt].BatchQty;
                            DI_oReturns.Lines.LineTotal = itemInfoList[i].LineTotal;

                            if (BatchCnt < i + RowCnt)
                            {
                                DI_oReturns.Lines.BatchNumbers.Add();
                            }
                        }
                        i += RowCnt -1;
                    }
                    else
                    {
                        DI_oReturns.Lines.Quantity = itemInfoList[i].OpenQty;
                        DI_oReturns.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                        DI_oReturns.Lines.LineTotal = itemInfoList[i].LineTotal;
                    }
                    DI_oReturns.Lines.Add();

                    //납품_S
                    DI_oDeliveryNotes.CardCode = itemInfoList[j].CardCode.Trim();
                    DI_oDeliveryNotes.DocDate = ODLNDocDate;
                    DI_oDeliveryNotes.DocDueDate = ODLNDocDate;
                    DI_oDeliveryNotes.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());
                    DI_oDeliveryNotes.Comments = "이월반품처리 [판매오더 : " + itemInfoList[j].ODLNDocB.Trim() + "]";

                    //헤더 환율처리_S (2011.11.04 송명규 추가)
                    if (lclDocCur != "KRW")
                    {
                        DI_oDeliveryNotes.DocCurrency = lclDocCur;
                        DI_oDeliveryNotes.DocRate = lclDocRate;
                    }
                    //헤더 환율처리_E

                    DI_oDeliveryNotes.Lines.ItemCode = itemInfoList[j].ItemCode;
                    DI_oDeliveryNotes.Lines.WarehouseCode = itemInfoList[j].WhsCode;
                    DI_oDeliveryNotes.Lines.UnitPrice = itemInfoList[j].Price;
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseType").Value = itemInfoList[j].BaseType; //기준납품문서의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[j].BaseEntry; //기준납품문서의 납품처리[SD404]문서번호의 값을 그대로 저장(2016.10.20 송명규 수정)
                    DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[j].BaseLine; //기준납품문서의 납품처리[SD404]문서행번호의 값을 그대로 저장(2016.10.20 송명규 수정)

                    //라인 환율처리_S
                    if (lclDocCur != "KRW")
                    {
                        DI_oDeliveryNotes.Lines.Currency = lclDocCur;
                        DI_oDeliveryNotes.Lines.Rate = lclDocRate;
                    }
                    //라인 환율처리_E

                    if (!string.IsNullOrEmpty(itemInfoList[j].BatchNum))
                    {
                        DI_oDeliveryNotes.Lines.Quantity = itemInfoList[j].OpenQty;
                        
                        for (int BatchCnt = j; BatchCnt < j + RowCnt; BatchCnt++)
                        {
                            DI_oDeliveryNotes.Lines.BatchNumbers.BatchNumber = itemInfoList[BatchCnt].BatchNum;
                            DI_oDeliveryNotes.Lines.BatchNumbers.Quantity = itemInfoList[BatchCnt].BatchQty;
                            DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[BatchCnt].BatchQty;
                            DI_oDeliveryNotes.Lines.LineTotal = itemInfoList[i].LineTotal;
                            if (BatchCnt < j + RowCnt)
                            {
                                DI_oDeliveryNotes.Lines.BatchNumbers.Add();
                            }
                            subRowCnt++;
                        }
                        j += RowCnt - 1;
                    }
                    else
                    {
                        DI_oDeliveryNotes.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[j].Qty;
                        DI_oDeliveryNotes.Lines.Quantity = itemInfoList[j].OpenQty;
                        DI_oDeliveryNotes.Lines.LineTotal = itemInfoList[j].LineTotal;
                        subRowCnt++;
                    }
                    DI_oDeliveryNotes.Lines.Add();
                    
                    if (itemInfoList.Count - 1 == i || itemInfoList[i].ODLNDocB != itemInfoList[i + 1].ODLNDocB) // 마지막행, 납품문서가 다를 경우 아래 구문실행
                    {
                        if (DI_oReturns != null) //반품 완료
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

                        if (DI_oDeliveryNotes != null) //납품 완료
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

                        k = subRowCnt;
                        for (lineCnt = 0; lineCnt < subRowCnt; lineCnt++, k--)
                        {
                            if (lineCnt != 0 && itemInfoList[i - k].ODLNDocB +"_"+ itemInfoList[i - k].DLN1LinB != itemInfoList[i - k + 1].ODLNDocB + "_" + itemInfoList[i - k + 1].DLN1LinB) //첫번째 행 외 납품문서 && 라인번호가 다르면 1플러스 됨
                            {
                                l++;
                            }
                            ItemInformation2 itemInfo2 = new ItemInformation2
                            {
                                ODLNDocB = itemInfoList[i - k + 1].ODLNDocB,
                                DLN1LinB = itemInfoList[i - k + 1].DLN1LinB,
                                ORDNDocEnt = ORDNDocEntry,
                                RDN1Lin = Convert.ToString(l),
                                ODLNDocEnt = ODLNDocEntry,
                                DLN1Lin = Convert.ToString(l)
                            };
                            itemInfoList2.Add(itemInfo2); 
                        }
                        l = 0;
                        subRowCnt = 0;
                        DI_oReturns = null; // 신규 납품,반품 생성시엔 초기화해야 오류 나지 않음.
                        DI_oDeliveryNotes = null;
                    }
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + itemInfoList.Count + "건의 납품 반품 문서 생성중...!";
                }
                oMat01.LoadFromDataSource();
                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
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
        /// SD920L 반품, 납품 문서번호 업데이트
        /// </summary>
        private bool PS_SD920_UpdateData()
        {
            bool returnValue = false;
            int i;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                for (i = 0; i < itemInfoList2.Count; i++)
                {
                    sQry = "update [@PS_SD920L] set U_ORDNDoc ='" + itemInfoList2[i].ORDNDocEnt + "',U_RDN1Line='" + itemInfoList2[i].RDN1Lin + "',U_ODLNDoc='" + itemInfoList2[i].ODLNDocEnt + "',U_DLN1Line='" + itemInfoList2[i].DLN1Lin + "', U_Check ='Y'";
                    sQry += "where DocEntry = '"+ DocEntry +"'and U_ODLNDocB ='" + itemInfoList2[i].ODLNDocB + "' and U_DLN1LinB ='" + itemInfoList2[i].DLN1LinB +"'";
                    oRecordSet01.DoQuery(sQry);
                }
                itemInfoList2.Clear();
                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                oForm.Freeze(true);
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
                                DocEntry = oForm.Items.Item("DocEntry").Specific.value;
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
                            if (PS_SD920_UpdateData() == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("납품, 반품 문서 업데이트중 오류발생(전산팀 연락바랍니다)");
                                BubbleEvent = false;
                            }
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                        PS_SD920_EnableFormItem();
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
                oForm.Freeze(false);
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
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //방향키(↑)
                        {
                            if (pVal.Row > 1 && pVal.Row <= oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.CharPressed == 40) //방향키(↓)
                        {
                            if (pVal.Row > 0 && pVal.Row < oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
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
                            PS_SD920_EnableFormItem();
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
                            PS_SD920_EnableFormItem();
                            break;
                        case "1282": //추가
                            PS_SD920_EnableFormItem();
                            PS_SD920_SetInitial();
                            PS_SD920_SetDocEntry();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD920_EnableFormItem();
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