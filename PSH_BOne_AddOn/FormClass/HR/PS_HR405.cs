using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전문직 인사평가 참고사항
    /// </summary>
    internal class PS_HR405 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_HR405H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_HR405L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocEntry01;
        private string gYear;
        private string gMSTCOD;
        private string gFULLNAME;
        private string gEmpNo1;
        private string gEmpName1;
        private string gPassWd;

        private SAPbouiCOM.BoFormMode oLast_Mode;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR405.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_HR405_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_HR405");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_HR405_CreateItems();
                PS_HR405_ComboBox_Setting();
                PS_HR405_SetDocument(oFromDocEntry01);
                PS_HR405_AddMatrixRow(0, true);
                PS_HR405_LoadCaption();

                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), false); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1285"), false); // 복원
                oForm.EnableMenu(("1284"), true); // 취소
                oForm.EnableMenu(("1293"), true); // 행삭제
                oForm.EnableMenu(("1281"), false);
                oForm.EnableMenu(("1282"), true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_HR405_CreateItems()
        {
            try
            {
                oDS_PS_HR405H = oForm.DataSources.DBDataSources.Item("@PS_HR405H");
                oDS_PS_HR405L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("PassWd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PassWd").Specific.DataBind.SetBound(true, "", "PassWd");

                oForm.Items.Item("Year").Specific.VALUE = DateTime.Now.ToString("yyyy");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_HR405_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_HR405_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "년도는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FULLNAME").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가자명은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "피평가자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("EmpName1").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "피평가자명은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("RateCode").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가항목은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("RateMNm").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가항목명은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가일은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Contents").Specific.VALUE.ToString().Trim()))
                {
                    errMessage = "평가내용은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_HR405_SetDocument(string oFromDocEntry01)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]";
                oRecordSet01.DoQuery(sQry);
                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_HR405_FormReset()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }
                oDS_PS_HR405H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
                oDS_PS_HR405H.SetValue("U_Year", 0, gYear);
                oDS_PS_HR405H.SetValue("U_MSTCOD", 0, gMSTCOD);
                oDS_PS_HR405H.SetValue("U_FULLNAME", 0, gFULLNAME);
                oDS_PS_HR405H.SetValue("U_EmpNo1", 0, "");
                oDS_PS_HR405H.SetValue("U_EmpName1", 0, "");
                oDS_PS_HR405H.SetValue("U_RateCode", 0, "");
                oDS_PS_HR405H.SetValue("U_RateMNm", 0, "");
                oDS_PS_HR405H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PS_HR405H.SetValue("U_Contents", 0, "");
                oDS_PS_HR405H.SetValue("U_EvalScr", 0, Convert.ToString(0));
                oForm.Items.Item("PassWd").Specific.VALUE = gPassWd;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_HR405_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_HR405_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_HR405L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_HR405L.Offset = oRow;
                oDS_PS_HR405L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PS_HR405_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "MSTCOD": //평가자
                        sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim() + "'";
                        sQry += " And  branch = '" + oForm.Items.Item("BPLId").Specific.VALUE + "' ";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("FULLNAME").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "EmpNo1": //피평가자
                        sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim() + "'";
                        sQry += " And  branch = '" + oForm.Items.Item("BPLId").Specific.VALUE + "' ";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("EmpName1").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "RateCode": //평가항목
                        sQry = "SELECT  T1.U_RateMNm FROM [@PS_HR400H] AS T0 Inner Join [@PS_HR400L] AS T1 On T0.Code = T1.Code ";
                        sQry += " WHERE   T0.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.VALUE + "' ";
                        sQry += " And T0.U_Year = '" + oForm.Items.Item("Year").Specific.VALUE + "' ";
                        sQry += " And T1.U_RateCode = '" + oForm.Items.Item("RateCode").Specific.VALUE + "'";

                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("RateMNm").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// LoadCaption
        /// </summary>
        private void PS_HR405_LoadCaption()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Btn_save").Specific.Caption = "추가";
                    oForm.Items.Item("Btn_del").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("Btn_save").Specific.Caption = "수정";
                    oForm.Items.Item("Btn_del").Enabled = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_HR405_LoadData
        /// </summary>
        private void PS_HR405_LoadData()
        {
            int i;
            string sQry;
            string SMSTCOD;
            string SBPLID;
            string SYear;
            string SEmpNo1;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                SBPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                SYear = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                SMSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                SEmpNo1 = oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim();

                if (string.IsNullOrEmpty(SEmpNo1))
                {
                    SEmpNo1 = "%";
                }

                sQry = "EXEC [PS_HR405_01] '" + SBPLID + "','" + SYear + "', '" + SMSTCOD + "', '" + SEmpNo1 + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_HR405L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    PS_HR405_AddMatrixRow(0, true);
                    PS_HR405_LoadCaption();
                    throw new Exception();
                }

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_HR405L.Size)
                    {
                        oDS_PS_HR405L.InsertRecord(i);
                    }
                    oMat01.AddRow();
                    oDS_PS_HR405L.Offset = i;
                    oDS_PS_HR405L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_HR405L.SetValue("U_ColReg01", i,oRecordSet01.Fields.Item(0).Value.ToString().Trim()); //BPLId
                    oDS_PS_HR405L.SetValue("U_ColReg02", i,oRecordSet01.Fields.Item(1).Value.ToString().Trim()); //DocEntry
                    oDS_PS_HR405L.SetValue("U_ColReg03", i,oRecordSet01.Fields.Item(2).Value.ToString().Trim()); //Year
                    oDS_PS_HR405L.SetValue("U_ColReg04", i,oRecordSet01.Fields.Item(3).Value.ToString().Trim()); //FULLNAME
                    oDS_PS_HR405L.SetValue("U_ColReg05", i,oRecordSet01.Fields.Item(4).Value.ToString().Trim()); //EmpName1
                    oDS_PS_HR405L.SetValue("U_ColReg06", i,oRecordSet01.Fields.Item(5).Value.ToString().Trim()); //RateCode
                    oDS_PS_HR405L.SetValue("U_ColReg07", i,oRecordSet01.Fields.Item(6).Value.ToString().Trim()); //RateName
                    oDS_PS_HR405L.SetValue("U_ColReg08", i,oRecordSet01.Fields.Item(7).Value.ToString().Trim()); //DocDate
                    oDS_PS_HR405L.SetValue("U_ColReg09", i,oRecordSet01.Fields.Item(8).Value.ToString().Trim()); //Contents
                    oDS_PS_HR405L.SetValue("U_ColSum01", i,oRecordSet01.Fields.Item(9).Value.ToString().Trim()); //EvalScr

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                if(ProgBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DeleteData
        /// </summary>
        private void PS_HR405_DeleteData()
        {
            string sQry;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    gYear = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                    gMSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                    gFULLNAME = oForm.Items.Item("FULLNAME").Specific.VALUE.ToString().Trim();
                    gEmpNo1 = oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim();
                    gEmpName1 = oForm.Items.Item("EmpName1").Specific.VALUE.ToString().Trim();
                    gPassWd = oForm.Items.Item("PassWd").Specific.VALUE.ToString().Trim();
                    DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();

                    sQry = "Select Count(*) From [@PS_HR405H] where DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.RecordCount == 0)
                    {
                        errMessage = "삭제대상이 없습니다. 확인하세요.";
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "Delete From [@PS_HR405H] where DocEntry = '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                PS_HR405_FormReset();
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// UpdateData
        /// </summary>
        /// <returns></returns>
        private bool PS_HR405_UpdateData(SAPbouiCOM.ItemEvent pVal)
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry;
            string RateMNm;
            string EmpName1;
            string FULLNAME;
            string Year_Renamed;
            string DocEntry;
            string BPLID;
            string MSTCOD;
            string EmpNo1;
            string RateCode;
            string DocDate;
            string Contents;
            string EvalScr;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                FULLNAME = oForm.Items.Item("FULLNAME").Specific.VALUE.ToString().Trim();
                EmpNo1 = oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim();
                EmpName1 = oForm.Items.Item("EmpName1").Specific.VALUE.ToString().Trim();
                RateCode = oForm.Items.Item("RateCode").Specific.VALUE.ToString().Trim();
                RateMNm = oForm.Items.Item("RateMNm").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                Contents = oForm.Items.Item("Contents").Specific.VALUE.ToString().Trim();
                EvalScr = oForm.Items.Item("EvalScr").Specific.VALUE.ToString().Trim();

                if (string.IsNullOrEmpty(DocEntry.ToString().Trim()))
                {
                    errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!";
                    throw new Exception();
                }

                sQry = "Update [@PS_HR405H]";
                sQry += " set ";
                sQry += " U_BPLId = '" + BPLID + "',";
                sQry += " U_Year = '" + Year_Renamed + "',";
                sQry += " U_MSTCOD = '" + MSTCOD + "',";
                sQry += " U_FULLNAME = '" + FULLNAME + "',";
                sQry += " U_Empno1 = '" + EmpNo1 + "',";
                sQry += " U_EmpName1 = '" + EmpName1 + "',";
                sQry += " U_RateCode  = '" + RateCode + "',";
                sQry += " U_RateMNm  = '" + RateMNm + "',";
                sQry += " U_DocDate  = '" + DocDate + "',";
                sQry += " U_Contents = '" + Contents + "',";
                sQry += " U_EvalScr = '" + EvalScr + "'";
                sQry += " Where DocEntry = '" + DocEntry + "'";
                oRecordSet01.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("수정완료", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success);

                gYear = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                gMSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                gFULLNAME = oForm.Items.Item("FULLNAME").Specific.VALUE.ToString().Trim();
                gEmpNo1 = oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim();
                gEmpName1 = oForm.Items.Item("EmpName1").Specific.VALUE.ToString().Trim();
                gPassWd = oForm.Items.Item("PassWd").Specific.VALUE.ToString().Trim();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else if (type == "M")
                    {
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PS_HR405_Add_PurchaseDemand
        /// </summary>
        /// <returns></returns>
        private bool PS_HR405_Add_PurchaseDemand(SAPbouiCOM.ItemEvent pVal)
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry;
            string RateMNm;
            string EmpName1;
            string FULLNAME;
            string Year_Renamed;
            string DocEntry;
            string BPLID;
            string MSTCOD;
            string EmpNo1;
            string RateCode;
            string DocDate;
            string Contents;
            string EvalScr;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                FULLNAME = oForm.Items.Item("FULLNAME").Specific.VALUE.ToString().Trim();
                EmpNo1 = oForm.Items.Item("EmpNo1").Specific.VALUE.ToString().Trim();
                EmpName1 = oForm.Items.Item("EmpName1").Specific.VALUE.ToString().Trim();
                RateCode = oForm.Items.Item("RateCode").Specific.VALUE.ToString().Trim();
                RateMNm = oForm.Items.Item("RateMNm").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                Contents = oForm.Items.Item("Contents").Specific.VALUE.ToString().Trim();
                EvalScr = oForm.Items.Item("EvalScr").Specific.VALUE.ToString().Trim();

                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]";
                oRecordSet01.DoQuery(sQry);
                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = Convert.ToString(1);
                }
                else
                {
                    DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1);
                }
                if (PS_HR405_PasswordChk(pVal) == false)
                {
                    errMessage = "패스워드가 틀렸습니다. 확인바랍니다.";
                    oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    sQry = "INSERT INTO [@PS_HR405H]";
                    sQry += " (";
                    sQry += " DocEntry,";
                    sQry += " DocNum,";
                    sQry += " U_BPLId,";
                    sQry += " U_Year,";
                    sQry += " U_MSTCOD,";
                    sQry += " U_FULLNAME,";
                    sQry += " U_EmpNo1,";
                    sQry += " U_EmpName1,";
                    sQry += " U_RateCode,";
                    sQry += " U_RateMNm,";
                    sQry += " U_DocDate,";
                    sQry += " U_Contents,";
                    sQry += " U_EvalScr";
                    sQry += " ) ";
                    sQry += "VALUES(";
                    sQry += DocEntry + ",";
                    sQry += DocEntry + ",";
                    sQry += "'" + BPLID + "',";
                    sQry += "'" + Year_Renamed + "',";
                    sQry += "'" + MSTCOD + "',";
                    sQry += "'" + FULLNAME + "',";
                    sQry += "'" + EmpNo1 + "',";
                    sQry += "'" + EmpName1 + "',";
                    sQry += "'" + RateCode + "',";
                    sQry += "'" + RateMNm + "',";
                    sQry += "'" + DocDate + "',";
                    sQry += "'" + Contents + "',";
                    sQry += "'" + EvalScr + "'";
                    sQry += ")";
                    oRecordSet02.DoQuery(sQry);

                    PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success);

                    gYear = Year_Renamed;
                    gMSTCOD = MSTCOD;
                    gFULLNAME = FULLNAME;
                    gEmpNo1 = EmpNo1;
                    gEmpName1 = EmpName1;
                    gPassWd = oForm.Items.Item("PassWd").Specific.VALUE.ToString().Trim();
                    PS_HR405_FormReset();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else if (type == "M")
                    {
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PasswordChk
        /// </summary>
        /// <returns></returns>
        private bool PS_HR405_PasswordChk(SAPbouiCOM.ItemEvent pVal)
        {
            bool returnValue = false;
            string sQry;
            string MSTCOD;
            string PassWd;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                PassWd = oForm.Items.Item("PassWd").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(MSTCOD.ToString().Trim()))
                {
                    errMessage = "사번이 없습니다. 입력바랍니다.";
                    throw new Exception();
                }

                sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                sQry += " And  BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "' ";
                sQry += " And  PassWd = '" + oForm.Items.Item("PassWd").Specific.Value + "' ";
                RecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) <= 0)
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
            return returnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_save")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_HR405_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_HR405_Add_PurchaseDemand(pVal) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_HR405_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_HR405_LoadCaption();
                            PS_HR405_LoadData();
                            oLast_Mode = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_HR405_UpdateData(pVal) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_HR405_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_HR405_LoadCaption();
                            PS_HR405_LoadData();
                        }
                    }
                    else if (pVal.ItemUID == "Btn_ret")
                    {
                        if (PS_HR405_PasswordChk(pVal) == false)
                        {
                            PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                            oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            PS_HR405_LoadData();
                        }
                    }
                    else if (pVal.ItemUID == "Btn_del")
                    {
                        PS_HR405_DeleteData();
                        PS_HR405_LoadData();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    if (pVal.ItemUID == "EmpNo1")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("EmpNo1").Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }

                    }
                    if (pVal.ItemUID == "RateCode")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("RateCode").Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            string DocEntry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);

                            DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE;

                            if (!string.IsNullOrEmpty(DocEntry.ToString().Trim()))
                            {
                                sQry = "EXEC [PS_HR405_02] '" + DocEntry + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                                    PS_HR405_LoadCaption();
                                    return;
                                }

                                oDS_PS_HR405H.SetValue("DocEntry", 0, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_BPLId", 0, oRecordSet01.Fields.Item("BPLId").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_Year", 0, oRecordSet01.Fields.Item("Year").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_MSTCOD", 0, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_FULLNAME", 0, oRecordSet01.Fields.Item("FULLNAME").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_EmpNo1", 0, oRecordSet01.Fields.Item("EmpNo1").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_EmpName1", 0, oRecordSet01.Fields.Item("EmpName1").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_RateCode", 0, oRecordSet01.Fields.Item("RateCode").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_RateMNm", 0, oRecordSet01.Fields.Item("RateMNm").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_DocDate", 0, oRecordSet01.Fields.Item("DocDate").Value.ToString("yyyyMMdd").Trim());
                                oDS_PS_HR405H.SetValue("U_Contents", 0, oRecordSet01.Fields.Item("Contents").Value.ToString().Trim());
                                oDS_PS_HR405H.SetValue("U_EvalScr", 0, oRecordSet01.Fields.Item("EvalScr").Value.ToString().Trim());
                                
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                PS_HR405_LoadCaption();
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "MSTCOD")
                        {
                            PS_HR405_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "EmpNo1")
                        {
                            PS_HR405_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "RateCode")
                        {
                            PS_HR405_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "PassWd")
                        {
                            if (PS_HR405_PasswordChk(pVal) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                                oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_HR405H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_HR405L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            PS_HR405_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PS_HR405_LoadCaption();
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
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
                            break;
                        case "1281": //찾기
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
