
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사용외출등록
    /// </summary>
    internal class PH_PY032 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY032A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY032B;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Mode;      //마지막 모드
        ////사용자구조체
        private struct ItemInformations
        {
            public string ItemCode;
            public string LotNo;
            public int Quantity;
            public int OPORNo;
            public int POR1No;
            public bool check;
            public int OPDNNo;
            public int PDN1No;
        }

        private ItemInformations[] ItemInformation;
        private int ItemInformationCount;

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY032.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY032_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY032");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                //oForm.DataBrowser.BrowseBy = "Code";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY032_CreateItems();
                PH_PY032_ComboBox_Setting();
                PH_PY032_EnableMenus();
                PH_PY032_SetDocument(oFromDocEntry01);
                PH_PY032_FormResize();

                PH_PY032_LoadCaption();
                PH_PY032_FormItemEnabled();

                oForm.EnableMenu(("1283"), false);                //// 삭제
                oForm.EnableMenu(("1286"), false);                //// 닫기
                oForm.EnableMenu(("1287"), false);                //// 복제
                oForm.EnableMenu(("1285"), false);                //// 복원
                oForm.EnableMenu(("1284"), false);                //// 취소
                oForm.EnableMenu(("1293"), false);                //// 행삭제
                oForm.EnableMenu(("1281"), false);
                oForm.EnableMenu(("1282"), true);

                string sQry = string.Empty;
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]";
                oRecordSet.DoQuery(sQry);
                if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(1));
                }
                else
                {
                    oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
                }
                PH_PY032_FormReset();
                //폼초기화 추가(2013.01.29 송명규)
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <returns></returns>
        private void PH_PY032_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY032A = oForm.DataSources.DBDataSources.Item("@PH_PY032A");
                oDS_PH_PY032B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //// 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //관리번호
                oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

                //사업장
                oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

                //부서
                oForm.DataSources.UserDataSources.Add("STeamCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STeamCd").Specific.DataBind.SetBound(true, "", "STeamCd");

                //사원번호
                oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

                //사원성명
                oForm.DataSources.UserDataSources.Add("SMSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SMSTNAM").Specific.DataBind.SetBound(true, "", "SMSTNAM");

                //시작월
                oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

                //종료월
                oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

                //목적
                oForm.DataSources.UserDataSources.Add("SObject", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
                oForm.Items.Item("SObject").Specific.DataBind.SetBound(true, "", "SObject");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY032_LoadCaption()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    oForm.Items.Item("BtnDelete").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    oForm.Items.Item("BtnDelete").Enabled = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_LoadCaption_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY032_Add_MatrixRow
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY032_Add_MatrixRow(int oRow, bool RowIserted = false)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (RowIserted == false)
                {
                    oDS_PH_PY032B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY032B.Offset = oRow;
                oDS_PH_PY032B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY032_MTX01
        /// </summary>
        private void PH_PY032_MTX01()
        {
            short i = 0;
            string sQry = string.Empty;
            string errCode = string.Empty;
            string sDocEntry = string.Empty;            //관리번호
            string sCLTCOD = string.Empty;            //사업장
            string sTeamCd = string.Empty;
            string sMSTCOD = string.Empty;            //사원번호
            string SFrDate = string.Empty;            //시작일자
            string SToDate = string.Empty;            //종료일자
            string SObject = string.Empty;            //목적

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);
            try
            {
                sCLTCOD = oForm.Items.Item("SCLTCOD").Specific.VALUE.ToString().Trim();                    //사업장
                sTeamCd = oForm.Items.Item("STeamCd").Specific.VALUE.ToString().Trim();                    //부서
                sMSTCOD = oForm.Items.Item("SMSTCOD").Specific.VALUE.ToString().Trim();                    //사원번호
                SFrDate = oForm.Items.Item("SFrDate").Specific.VALUE.ToString().Trim().Replace(".", "");   //시작일자
                SToDate = oForm.Items.Item("SToDate").Specific.VALUE.ToString().Trim().Replace(".", "");   //종료일자
                SObject = oForm.Items.Item("SObject").Specific.VALUE.ToString().Trim();                    //목적

                oForm.Freeze(true);

                sQry = "                EXEC [PH_PY032_01] ";
                sQry = sQry + "'" + sDocEntry + "',";     //관리번호
                sQry = sQry + "'" + sCLTCOD + "',";       //사업장
                sQry = sQry + "'" + sTeamCd + "',";       //부서
                sQry = sQry + "'" + sMSTCOD + "',";       //사원번호
                sQry = sQry + "'" + SFrDate + "',";       //시작일자
                sQry = sQry + "'" + SToDate + "',";       //종료일자
                sQry = sQry + "'" + SObject + "'";        //목적

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY032B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if ((oRecordSet01.RecordCount == 0))
                {
                    errCode = "1";
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    PH_PY032_LoadCaption();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY032B.Size)
                    {
                        oDS_PH_PY032B.InsertRecord((i));
                    }
                    oMat01.AddRow();
                    oDS_PH_PY032B.Offset = i;

                    oDS_PH_PY032B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY032B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());                    //관리번호
                    oDS_PH_PY032B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim());                    //사업장
                    oDS_PH_PY032B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());                    //사원번호
                    oDS_PH_PY032B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim());                    //사원성명
                    oDS_PH_PY032B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("FrDate").Value.ToString().Trim());                    //시작일자
                    oDS_PH_PY032B.SetValue("U_ColTm01", i, oRecordSet01.Fields.Item("FrTime").Value.ToString().Trim());                    //시작시각
                    oDS_PH_PY032B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("ToDate").Value.ToString().Trim());                    //종료일자
                    oDS_PH_PY032B.SetValue("U_ColTm02", i, oRecordSet01.Fields.Item("ToTime").Value.ToString().Trim());                    //종료시각
                    oDS_PH_PY032B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("Object").Value.ToString().Trim());                    //목적

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY032_DeleteData
        /// </summary>
        private void PH_PY032_DeleteData()
        {
            short i = 0;
            string sQry = string.Empty; ;
            short ErrNum = 0;
            string DocEntry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PH_PY032A] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if ((oRecordSet01.RecordCount == 0))
                    {
                        ErrNum = 1;
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "EXEC PH_PY032_04 '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                dataHelpClass.MDC_GF_Message("삭제 완료!", "S");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_DeleteData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY032_UpdateData
        /// </summary>
        /// <returns></returns>
        private bool PH_PY032_UpdateData()
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            int DocEntry = 0;
            string CLTCOD = string.Empty;                //사업장
            string MSTCOD = string.Empty;                //사원번호
            string MSTNAM = string.Empty;                //사원성명
            string FrDate = string.Empty;                //시작일자
            string FrTime = string.Empty;                //시작시각
            string ToDate = string.Empty;                //종료일자
            string ToTime = string.Empty;                //종료시각
            string Object_Renamed = string.Empty;        //목적


            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim());       //관리번호
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();                            //사업장
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();                            //사원번호
                MSTNAM = oForm.Items.Item("MSTNAM").Specific.VALUE.ToString().Trim();                            //사원성명
                FrDate = oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim();                            //시작일자
                FrTime = oForm.Items.Item("FrTime").Specific.VALUE.ToString().Trim();                            //시작시각
                ToDate = oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim();                            //종료일자
                ToTime = oForm.Items.Item("ToTime").Specific.VALUE.ToString().Trim();                            //종료시각
                Object_Renamed = oForm.Items.Item("Object").Specific.VALUE.ToString().Trim();                    //목적

                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).ToString().Trim()))
                {
                    dataHelpClass.MDC_GF_Message("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E");
                    functionReturnValue = false;
                    throw new Exception();
                }

                sQry = "                EXEC [PH_PY032_03] ";
                sQry = sQry + "'" + DocEntry + "',";              //관리번호
                sQry = sQry + "'" + CLTCOD + "',";                //사업장
                sQry = sQry + "'" + MSTCOD + "',";                //사원번호
                sQry = sQry + "'" + MSTNAM + "',";                //사원성명
                sQry = sQry + "'" + FrDate + "',";                //시작일자
                sQry = sQry + "'" + FrTime + "',";                //시작시각
                sQry = sQry + "'" + ToDate + "',";                //종료일자
                sQry = sQry + "'" + ToTime + "',";                //종료시각
                sQry = sQry + "'" + Object_Renamed + "'";         //목적

                oRecordSet01.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("수정 완료!", "S");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_UpdateData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }
            finally
            {
                functionReturnValue = true;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY032_DeleteData
        /// </summary>
        private bool PH_PY032_AddData()
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            int DocEntry = 0;
            string CLTCOD = string.Empty;                //사업장
            string MSTCOD = string.Empty;                //사원번호
            string MSTNAM = string.Empty;                //사원성명
            string FrDate = string.Empty;                //시작일자
            string FrTime = string.Empty;                //시작시각
            string ToDate = string.Empty;                //종료일자
            string ToTime = string.Empty;                //종료시각
            string Object_Renamed = string.Empty;        //목적
            string UserSign = string.Empty;              //UserSign

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();                //사업장
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();                //사원번호
                MSTNAM = oForm.Items.Item("MSTNAM").Specific.VALUE.ToString().Trim();                //사원성명
                FrDate = oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim();                //시작일자
                FrTime = oForm.Items.Item("FrTime").Specific.VALUE.ToString().Trim();                //시작시각
                ToDate = oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim();                //종료일자
                ToTime = oForm.Items.Item("ToTime").Specific.VALUE.ToString().Trim();                //종료시각
                Object_Renamed = oForm.Items.Item("Object").Specific.VALUE.ToString().Trim();        //목적
                UserSign = PSH_Globals.oCompany.UserSignature.ToString();

                //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY032A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }

                sQry = "   EXEC [PH_PY032_02] ";
                sQry = sQry + "'" + DocEntry + "',";               //관리번호
                sQry = sQry + "'" + CLTCOD + "',";                //사업장
                sQry = sQry + "'" + MSTCOD + "',";                //사원번호
                sQry = sQry + "'" + MSTNAM + "',";                //사원성명
                sQry = sQry + "'" + FrDate + "',";                //시작일자
                sQry = sQry + "'" + FrTime + "',";                //시작시각
                sQry = sQry + "'" + ToDate + "',";                //종료일자
                sQry = sQry + "'" + ToTime + "',";                //종료시각
                sQry = sQry + "'" + Object_Renamed + "',";        //목적
                sQry = sQry + "'" + UserSign + "'";               //UserSign

                oRecordSet02.DoQuery(sQry);

                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_AddData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }
            finally
            {
                functionReturnValue = true;
            }
            return functionReturnValue;
        }


        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PH_PY032_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short ErrNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Items.Item("MSTCOD").Specific.Value.Trim() == "") //사원번호
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                else if (oForm.Items.Item("FrDate").Specific.Value.Trim() == "") //시작일자
                {
                    ErrNum = 4;
                    throw new Exception();
                }
                else if (oForm.Items.Item("FrTime").Specific.Value.Trim() == "") //시작시각
                {
                    ErrNum = 5;
                    throw new Exception();
                }
                else if (oForm.Items.Item("ToDate").Specific.Value == "") //종료일자
                {
                    ErrNum = 6;
                    throw new Exception();
                }
                else if (oForm.Items.Item("ToTime").Specific.Value == "") //종료시각
                {
                    ErrNum = 7;
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 3)
                {
                    dataHelpClass.MDC_GF_Message("사원번호는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 4)
                {
                    dataHelpClass.MDC_GF_Message("시작일자는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 5)
                {
                    dataHelpClass.MDC_GF_Message("시작시각은 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 6)
                {
                    dataHelpClass.MDC_GF_Message("종료일자는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 7)
                {
                    dataHelpClass.MDC_GF_Message("종료시각은 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_HeaderSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                functionReturnValue = false;
            }
            finally
            {
                functionReturnValue = true;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 메트릭스 필수 사항 check
        /// 구현은 되어 있지만 사용하지 않음
        /// </summary>
        /// <returns></returns>
        private bool PH_PY032_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;

            int i = 0;
            short ErrNum = 0;

            try
            {
                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                functionReturnValue = false;
            }

            return functionReturnValue;
        }


        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY032_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
        {
            int i = 0;
            short ErrNum = 0;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sCLTCOD = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                switch (oUID)
                {

                    case "SCLTCOD":

                        sCLTCOD = oForm.Items.Item("SCLTCOD").Specific.VALUE.ToString().Trim();
                        if (oForm.Items.Item("STeamCd").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("STeamCd").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("STeamCd").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }
                        //부서콤보세팅
                        oForm.Items.Item("STeamCd").Specific.ValidValues.Add("%", "전체");
                        sQry = "            SELECT      U_Code AS [Code],";
                        sQry = sQry + "                 U_CodeNm As [Name]";
                        sQry = sQry + "  FROM       [@PS_HR200L]";
                        sQry = sQry + "  WHERE      Code = '1'";
                        sQry = sQry + "                 AND U_UseYN = 'Y'";
                        sQry = sQry + "                 AND U_Char2 = '" + sCLTCOD + "'";
                        sQry = sQry + "  ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("STeamCd").Specific, sQry, "", false, false);
                        oForm.Items.Item("STeamCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY032_DeleteData
        /// </summary>
        public void PH_PY032_ComboBox_Setting()
        {
            short i = 0;
            string sQry = string.Empty; ;
            short ErrNum = 0;
            string DocEntry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                ////////////매트릭스//////////
                //사업장
                sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), sQry, "", "");
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_ComboBox_Setting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY032_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);               ////취소
                oForm.EnableMenu("1287", true);                ////복제
                oForm.EnableMenu("1293", false);               ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY032_SetDocument(string oFromDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY032_FormItemEnabled();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_FormResize
        /// </summary>
        private void PH_PY032_FormResize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_FormResize_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_FormReset
        /// </summary>
        public void PH_PY032_FormReset()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                //관리번호
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(1));
                }
                else
                {
                    oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1));
                }

                string User_BPLID = null;
                User_BPLID = dataHelpClass.User_BPLID();

                ////////////기준정보//////////
                oDS_PH_PY032A.SetValue("U_CLTCOD", 0, User_BPLID);                            //사업장
                oDS_PH_PY032A.SetValue("U_MSTCOD", 0, "");                                    //사원번호
                oDS_PH_PY032A.SetValue("U_MSTNAM", 0, "");                                    //사원성명
                oDS_PH_PY032A.SetValue("U_FrDate", 0, DateTime.Now.ToString("yyyyMMdd"));     //시작일자
                oDS_PH_PY032A.SetValue("U_FrTime", 0, "");                                    //시작시각
                oDS_PH_PY032A.SetValue("U_ToDate", 0, DateTime.Now.ToString("yyyyMMdd"));     //종료일자
                oDS_PH_PY032A.SetValue("U_ToTime", 0, "");                                    //종료시각
                oDS_PH_PY032A.SetValue("U_Object", 0, "");                                    //목적

                oForm.Items.Item("MSTCOD").Click();

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_CalculateTransExp
        /// </summary>
        private void PH_PY032_CalculateTransExp()
        {
            try
            {
                double FuelPrc = 0;                //유류단가
                double Distance = 0;                //거리
                double TransExp = 0;                //교통비

                FuelPrc = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.VALUE.ToString().Trim());
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.VALUE.ToString().Trim());

                TransExp = ((FuelPrc * Distance * 0.1) / 10) * 10;                //원단위 절사

                oForm.Items.Item("TransExp").Specific.VALUE = TransExp;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_CalculateTransExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_CalculateTransExp
        /// </summary>
        private void PH_PY032_CalculateTotalExp()
        {
            double TransExp = 0;                //교통비
            double DayExp = 0;                //일비
            double FoodExp = 0;                //식비
            double ParkExp = 0;                //주차비
            double TollExp = 0;                //도로비
            double TotalExp = 0;                //합계
            try
            {
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.VALUE.ToString().Trim());
                DayExp = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.VALUE.ToString().Trim());
                FoodExp = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.VALUE.ToString().Trim());
                ParkExp = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.VALUE.ToString().Trim());
                TollExp = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.VALUE.ToString().Trim());
                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp;

                oDS_PH_PY032A.SetValue("U_TotalExp", 0, Convert.ToString(TotalExp));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_CalculateTotalExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_GetDestNo
        /// </summary>
        private void PH_PY032_GetDestNo()
        {
            string FrDate = string.Empty;
            string CLTCOD = string.Empty;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                FrDate = oForm.Items.Item("FrDate").Specific.VALUE.ToString("yyyy").Trim();

                sQry = "EXEC PH_PY032_05 '" + CLTCOD + "', '" + FrDate + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("DestNo1").Specific.VALUE = FrDate;
                oForm.Items.Item("DestNo2").Specific.VALUE = oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_GetDestNo_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY032_CalculateFoodExp
        /// </summary>
        private void PH_PY032_CalculateFoodExp()
        {
            short ErrNum = 0;
            string sQry = null;
            string MSTCOD = null;            //사번
            short FoodNum = 0;            //식수
            double FoodPrc = 0;            //당일식비
            double FoodExp = 0;            //전체식비
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                //사번을 선택하지 않으면
                if (string.IsNullOrEmpty(oDS_PH_PY032A.GetValue("U_MSTCOD", 0).ToString().Trim()) & oDS_PH_PY032A.GetValue("U_FoodNum", 0).ToString().Trim() != "0")
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                MSTCOD = oDS_PH_PY032A.GetValue("U_MSTCOD", 0).ToString().Trim();                       //사번
                FoodNum = Convert.ToInt16(oDS_PH_PY032A.GetValue("U_FoodNum", 0).ToString().Trim());    //식수

                sQry = "            SELECT      T1.U_Num4 AS [FoodPrc]";
                sQry = sQry + "  FROM       [@PH_PY001A] AS T0";
                sQry = sQry + "                 LEFT JOIN";
                sQry = sQry + "                 [@PS_HR200L] AS T1";
                sQry = sQry + "                     ON T0.U_JIGCOD = T1.U_Code";
                sQry = sQry + "                     AND T1.Code = 'P232'";
                sQry = sQry + "                     AND T1.U_UseYN = 'Y'";
                sQry = sQry + "  WHERE      T0.Code = '" + MSTCOD + "'";

                oRecordSet01.DoQuery(sQry);

                FoodPrc = oRecordSet01.Fields.Item("FoodPrc").Value;
                FoodExp = FoodPrc * FoodNum;

                oDS_PH_PY032A.SetValue("U_FoodExp", 0, Convert.ToString(FoodExp));
                oForm.Items.Item("FoodExp").Click();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번을 선택하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_CalculateFoodExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY032_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY032_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //break;

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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    // break;

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
                    if (pVal.ItemUID == "PH_PY032")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }

                    ///추가/확인 버튼클릭
                    if (pVal.ItemUID == "BtnAdd")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY032_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY032_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY032_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY032_LoadCaption();
                            PH_PY032_MTX01();

                            oLast_Mode = (int)oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY032_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY032_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY032_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY032_LoadCaption();
                            PH_PY032_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch")
                    {
                        PH_PY032_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        PH_PY032_LoadCaption();
                        PH_PY032_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            PH_PY032_DeleteData();
                            PH_PY032_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY032_LoadCaption();
                            PH_PY032_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY032_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY032")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", "");                    //기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SMSTCOD", "");                    //조회조건-사번
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if ((pVal.ItemUID == "Mat01"))
                        {
                        }
                        else
                        {
                            PH_PY032_FlushToItemValue(pVal.ItemUID);
                            if (pVal.ItemUID == "MSTCOD")
                            {
                                oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'","");                                //성명
                            }
                            else if (pVal.ItemUID == "SMSTCOD")
                            {
                                oForm.Items.Item("SMSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.VALUE + "'", "");                                //성명
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
                oForm.Freeze(false);
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY032_FlushToItemValue(pVal.ItemUID);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY032_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY032_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// ROW_DELETE(Raise_FormMenuEvent에서 호출)
        /// 해당 클래스에서는 사용되지 않음
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
            try
            {
                if ((oLastColRow01 > 0))
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PH_PY032A.RemoveRecord(oDS_PH_PY032A.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY032_Add_MatrixRow(0);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY032A.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY032_Add_MatrixRow(oMat01.RowCount);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
                            oForm.Freeze(true);

                            //DataSource를 이용하여 각 컨트롤에 값을 출력
                            oDS_PH_PY032A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE);                          //관리번호
                            oDS_PH_PY032A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.VALUE);                            //사업장
                            oDS_PH_PY032A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE);                            //사원번호
                            oDS_PH_PY032A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.VALUE);                            //사원성명
                            oDS_PH_PY032A.SetValue("U_FrDate", 0, oMat01.Columns.Item("FrDate").Cells.Item(pVal.Row).Specific.VALUE.Replace(".", ""));           //시작일자
                            oDS_PH_PY032A.SetValue("U_FrTime", 0, oMat01.Columns.Item("FrTime").Cells.Item(pVal.Row).Specific.VALUE);                            //시작시각
                            oDS_PH_PY032A.SetValue("U_ToDate", 0, oMat01.Columns.Item("ToDate").Cells.Item(pVal.Row).Specific.VALUE.Replace(".", ""));           //종료일자
                            oDS_PH_PY032A.SetValue("U_ToTime", 0, oMat01.Columns.Item("ToTime").Cells.Item(pVal.Row).Specific.VALUE);                            //종료시각
                            oDS_PH_PY032A.SetValue("U_Object", 0, oMat01.Columns.Item("Object").Cells.Item(pVal.Row).Specific.VALUE);                            //목적

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PH_PY032_LoadCaption();

                            oForm.Freeze(false);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY032A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY032B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            break;
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            ///추가버튼 클릭시 메트릭스 insertrow

                            PH_PY032_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PH_PY032_LoadCaption();
                            return;
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;
                    }
                    ////BeforeAction = False
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":                            //취소
                            break;
                        case "1286":                            //닫기
                            break;
                        case "1293":                            //행삭제
                            break;
                        case "1281":                            //찾기
                            break;
                        case "1282":                            //추가
                            break;

                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i = 0;
            string sQry = string.Empty;

            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                    ////BeforeAction = False
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY032_Print_Report01()
        {
            int DocEntry = 0;
            string WinTitle = string.Empty;
            string ReportName = string.Empty;
            string CLTCOD = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            try
            {
                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim());
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                WinTitle = "[PH_PY032] 사용외출증";

                if (CLTCOD == "1")//창원
                {
                    ReportName = "PH_PY032_01.rpt";
                    
                }
                else if (CLTCOD == "2")//동래
                {
                    ReportName = "PH_PY032_02.rpt";
                    
                }
                else if (CLTCOD == "3")//사상
                {
                    ReportName = "PH_PY032_03.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                //dataPackFormula.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); //년도

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); //사업장


                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
               // formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY032_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}

//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
//// ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//    internal class PH_PY032
//    {
//        //****************************************************************************************************************
//        ////  File : PH_PY032.cls
//        ////  Module : 인사관리>기타관리
//        ////  Desc : 사용외출등록
//        ////  FormType : PH_PY032
//        ////  Create Date(Start) : 2013.03.19
//        ////  Create Date(End) :
//        ////  Creator : Song Myoung gyu
//        ////  Modified Date :
//        ////  Modifier :
//        ////  Company : Poongsan Holdings
//        //****************************************************************************************************************

//        public string oFormUniqueID01;
//        public SAPbouiCOM.Form oForm;
//        public SAPbouiCOM.Matrix oMat01;
//        //등록헤더
//        private SAPbouiCOM.DBDataSource oDS_PH_PY032A;
//        //등록라인
//        private SAPbouiCOM.DBDataSource oDS_PH_PY032B;

//        //클래스에서 선택한 마지막 아이템 Uid값
//        private string oLastItemUID01;
//        //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//        private string oLastColUID01;
//        //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//        private int oLastColRow01;

//        ////사용자구조체
//        private struct ItemInformations
//        {
//            public string ItemCode;
//            public string LotNo;
//            public int Quantity;
//            public int OPORNo;
//            public int POR1No;
//            public bool check;
//            public int OPDNNo;
//            public int PDN1No;
//        }

//        private int oLast_Mode;

//        private ItemInformations[] ItemInformation;
//        private int ItemInformationCount;

//        //*******************************************************************
//        // .srf 파일로부터 폼을 로드한다.
//        //*******************************************************************
//        public void LoadForm(string oFromDocEntry01 = "")
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            int i = 0;
//            string oInnerXml = null;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY032.srf");
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//            //매트릭스의 타이틀높이와 셀높이를 고정
//            for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//            {
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//            }

//            oFormUniqueID01 = "PH_PY032_" + GetTotalFormsCount();
//            SubMain.AddForms(this, oFormUniqueID01, "PH_PY032");
//            ////폼추가
//            MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//            //폼 할당
//            oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID01);

//            oForm.SupportedModes = -1;
//            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//            ////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

//            oForm.Freeze(true);
//            PH_PY032_CreateItems();
//            PH_PY032_ComboBox_Setting();
//            PH_PY032_CF_ChooseFromList();
//            PH_PY032_EnableMenus();
//            PH_PY032_SetDocument(oFromDocEntry01);
//            PH_PY032_FormResize();


//            //    Call PH_PY032_Add_MatrixRow(0, True)
//            PH_PY032_LoadCaption();
//            PH_PY032_FormItemEnabled();

//            oForm.EnableMenu(("1283"), false);
//            //// 삭제
//            oForm.EnableMenu(("1286"), false);
//            //// 닫기
//            oForm.EnableMenu(("1287"), false);
//            //// 복제
//            oForm.EnableMenu(("1285"), false);
//            //// 복원
//            oForm.EnableMenu(("1284"), false);
//            //// 취소
//            oForm.EnableMenu(("1293"), false);
//            //// 행삭제
//            oForm.EnableMenu(("1281"), false);
//            oForm.EnableMenu(("1282"), true);

//            string sQry = null;
//            SAPbobsCOM.Recordset RecordSet01 = null;
//            RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]";
//            RecordSet01.DoQuery(sQry);
//            if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0)
//            {
//                oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(1));
//            }
//            else
//            {
//                oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1));
//            }

//            PH_PY032_FormReset();
//            //폼초기화 추가(2013.01.29 송명규)

//            oForm.Update();
//            oForm.Freeze(false);

//            oForm.Visible = true;
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;

//            //기간(월)
//            //UPGRADE_WARNING: oForm.Items(SFrDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SFrDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY.MM");
//            //UPGRADE_WARNING: oForm.Items(SToDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SToDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY.MM");
//            //사번 포커스
//            oForm.Items.Item("MSTCOD").Click();

//            return;
//        LoadForm_Error:
//            oForm.Update();
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oForm = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY032_LoadCaption()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_LoadCaption()
//            //해당모듈 : PH_PY032
//            //기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
//                oForm.Items.Item("BtnDelete").Enabled = false;
//                //    ElseIf oForm.Mode = fm_OK_MODE Then
//                //        oForm.Items("BtnAdd").Specific.Caption = "확인"
//            }
//            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
//                oForm.Items.Item("BtnDelete").Enabled = true;
//            }

//            oForm.Freeze(false);

//            return;
//        PH_PY032_LoadCaption_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//        }

//        ///메트릭스 Row추가
//        public void PH_PY032_Add_MatrixRow(int oRow, ref bool RowIserted = false)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            ////행추가여부
//            if (RowIserted == false)
//            {
//                oDS_PH_PY032B.InsertRecord((oRow));
//            }

//            oMat01.AddRow();
//            oDS_PH_PY032B.Offset = oRow;
//            oDS_PH_PY032B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

//            oMat01.LoadFromDataSource();
//            return;
//        PH_PY032_Add_MatrixRow_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            MDC_Com.MDC_GF_Message(ref "PH_PY032_Add_MatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//        }

//        public void PH_PY032_MTX01()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_MTX01()
//            //해당모듈 : PH_PY032
//            //기능 : 데이터 조회
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short i = 0;
//            string sQry = null;
//            short ErrNum = 0;

//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            string sDocEntry = null;
//            //관리번호
//            string sCLTCOD = null;
//            //사업장
//            string sTeamCd = null;
//            string sMSTCOD = null;
//            //사원번호
//            string SFrDate = null;
//            //시작일자
//            string SToDate = null;
//            //종료일자
//            string SObject = null;
//            //목적

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            sCLTCOD = Strings.Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE);
//            //사업장
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            sTeamCd = Strings.Trim(oForm.Items.Item("STeamCd").Specific.VALUE);
//            //부서
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            sMSTCOD = Strings.Trim(oForm.Items.Item("SMSTCOD").Specific.VALUE);
//            //사원번호
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            SFrDate = Strings.Replace(Strings.Trim(oForm.Items.Item("SFrDate").Specific.VALUE), ".", "");
//            //시작일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            SToDate = Strings.Replace(Strings.Trim(oForm.Items.Item("SToDate").Specific.VALUE), ".", "");
//            //종료일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            SObject = Strings.Trim(oForm.Items.Item("SObject").Specific.VALUE);
//            //목적

//            SAPbouiCOM.ProgressBar ProgBar01 = null;
//            ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//            oForm.Freeze(true);

//            sQry = "                EXEC [PH_PY032_01] ";
//            sQry = sQry + "'" + sDocEntry + "',";
//            //관리번호
//            sQry = sQry + "'" + sCLTCOD + "',";
//            //사업장
//            sQry = sQry + "'" + sTeamCd + "',";
//            //부서
//            sQry = sQry + "'" + sMSTCOD + "',";
//            //사원번호
//            sQry = sQry + "'" + SFrDate + "',";
//            //시작일자
//            sQry = sQry + "'" + SToDate + "',";
//            //종료일자
//            sQry = sQry + "'" + SObject + "'";
//            //목적

//            oRecordSet01.DoQuery(sQry);

//            oMat01.Clear();
//            oDS_PH_PY032B.Clear();
//            oMat01.FlushToDataSource();
//            oMat01.LoadFromDataSource();

//            if ((oRecordSet01.RecordCount == 0))
//            {

//                ErrNum = 1;

//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                //        Call PH_PY032_Add_MatrixRow(0, True)
//                PH_PY032_LoadCaption();

//                goto PH_PY032_MTX01_Error;

//                return;
//            }

//            for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
//            {
//                if (i + 1 > oDS_PH_PY032B.Size)
//                {
//                    oDS_PH_PY032B.InsertRecord((i));
//                }

//                oMat01.AddRow();
//                oDS_PH_PY032B.Offset = i;

//                oDS_PH_PY032B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                oDS_PH_PY032B.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value));
//                //관리번호
//                oDS_PH_PY032B.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("CLTCOD").Value));
//                //사업장
//                oDS_PH_PY032B.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("MSTCOD").Value));
//                //사원번호
//                oDS_PH_PY032B.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("MSTNAM").Value));
//                //사원성명
//                oDS_PH_PY032B.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("FrDate").Value));
//                //시작일자
//                oDS_PH_PY032B.SetValue("U_ColTm01", i, Strings.Trim(oRecordSet01.Fields.Item("FrTime").Value));
//                //시작시각
//                oDS_PH_PY032B.SetValue("U_ColReg07", i, Strings.Trim(oRecordSet01.Fields.Item("ToDate").Value));
//                //종료일자
//                oDS_PH_PY032B.SetValue("U_ColTm02", i, Strings.Trim(oRecordSet01.Fields.Item("ToTime").Value));
//                //종료시각
//                oDS_PH_PY032B.SetValue("U_ColReg09", i, Strings.Trim(oRecordSet01.Fields.Item("Object").Value));
//                //목적

//                oRecordSet01.MoveNext();
//                ProgBar01.Value = ProgBar01.Value + 1;
//                ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

//            }

//            oMat01.LoadFromDataSource();
//            oMat01.AutoResizeColumns();
//            ProgBar01.Stop();
//            oForm.Freeze(false);

//            //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgBar01 = null;
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return;
//        PH_PY032_MTX01_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            //    ProgBar01.Stop
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgBar01 = null;
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            if (ErrNum == 1)
//            {
//                MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }
//        }

//        public void PH_PY032_DeleteData()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_DeleteData()
//            //해당모듈 : PH_PY032
//            //기능 : 기본정보 삭제
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short i = 0;
//            string sQry = null;
//            short ErrNum = 0;

//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            string DocEntry = null;

//            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//            {

//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);

//                sQry = "SELECT COUNT(*) FROM [@PH_PY032A] WHERE DocEntry = '" + DocEntry + "'";
//                oRecordSet01.DoQuery(sQry);

//                if ((oRecordSet01.RecordCount == 0))
//                {
//                    ErrNum = 1;
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                    goto PH_PY032_DeleteData_Error;
//                }
//                else
//                {
//                    sQry = "EXEC PH_PY032_04 '" + DocEntry + "'";
//                    oRecordSet01.DoQuery(sQry);
//                }
//            }

//            MDC_Com.MDC_GF_Message(ref "삭제 완료!", ref "S");

//            //    Call PH_PY032_FormReset

//            //    oForm.Mode = fm_ADD_MODE

//            //    Call oForm.Items("BtnSearch").Click(ct_Regular)

//            //    oMat01.Clear
//            //    oMat01.FlushToDataSource
//            //    oMat01.LoadFromDataSource
//            //    Call PH_PY032_Add_MatrixRow(0, True)

//            return;
//        PH_PY032_DeleteData_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            if (ErrNum == 1)
//            {
//                MDC_Com.MDC_GF_Message(ref "삭제대상이 없습니다. 확인하세요.", ref "W");
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_DeleteData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }
//        }

//        public bool PH_PY032_UpdateData()
//        {
//            bool functionReturnValue = false;
//            //******************************************************************************
//            //Function ID : PH_PY032_UpdateData()
//            //해당모듈 : PH_PY032
//            //기능 : 기본정보를 수정
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short i = 0;
//            short j = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset RecordSet01 = null;
//            RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            short DocEntry = 0;
//            //관리번호
//            string CLTCOD = null;
//            //사업장
//            string MSTCOD = null;
//            //사원번호
//            string MSTNAM = null;
//            //사원성명
//            string FrDate = null;
//            //시작일자
//            string FrTime = null;
//            //시작시각
//            string ToDate = null;
//            //종료일자
//            string ToTime = null;
//            //종료시각
//            //UPGRADE_NOTE: Object이(가) Object_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//            string Object_Renamed = null;
//            //목적

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = Convert.ToInt16(Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE));
//            //관리번호
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//            //사업장
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
//            //사원번호
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            MSTNAM = Strings.Trim(oForm.Items.Item("MSTNAM").Specific.VALUE);
//            //사원성명
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FrDate = Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE);
//            //시작일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FrTime = Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE);
//            //시작시각
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ToDate = Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE);
//            //종료일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ToTime = Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE);
//            //종료시각
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Object_Renamed = Strings.Trim(oForm.Items.Item("Object").Specific.VALUE);
//            //목적

//            if (string.IsNullOrEmpty(Strings.Trim(Convert.ToString(DocEntry))))
//            {
//                MDC_Com.MDC_GF_Message(ref "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", ref "E");
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            sQry = "                EXEC [PH_PY032_03] ";
//            sQry = sQry + "'" + DocEntry + "',";
//            //관리번호
//            sQry = sQry + "'" + CLTCOD + "',";
//            //사업장
//            sQry = sQry + "'" + MSTCOD + "',";
//            //사원번호
//            sQry = sQry + "'" + MSTNAM + "',";
//            //사원성명
//            sQry = sQry + "'" + FrDate + "',";
//            //시작일자
//            sQry = sQry + "'" + FrTime + "',";
//            //시작시각
//            sQry = sQry + "'" + ToDate + "',";
//            //종료일자
//            sQry = sQry + "'" + ToTime + "',";
//            //종료시각
//            sQry = sQry + "'" + Object_Renamed + "'";
//            //목적

//            RecordSet01.DoQuery(sQry);

//            MDC_Com.MDC_GF_Message(ref "수정 완료!", ref "S");

//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            functionReturnValue = true;
//            return functionReturnValue;
//        PH_PY032_UpdateData_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            MDC_Com.MDC_GF_Message(ref "PH_PY032_UpdateData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            return functionReturnValue;
//        }

//        public bool PH_PY032_AddData()
//        {
//            bool functionReturnValue = false;
//            //******************************************************************************
//            //Function ID : PH_PY032_AddData()
//            //해당모듈 : PH_PY032
//            //기능 : 데이터 INSERT
//            //인수 : 없음
//            //반환값 : 성공여부
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short i = 0;
//            string sQry = null;

//            SAPbobsCOM.Recordset RecordSet01 = null;
//            RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            SAPbobsCOM.Recordset RecordSet02 = null;
//            RecordSet02 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            short DocEntry = 0;
//            //관리번호
//            string CLTCOD = null;
//            //사업장
//            string MSTCOD = null;
//            //사원번호
//            string MSTNAM = null;
//            //사원성명
//            string FrDate = null;
//            //시작일자
//            string FrTime = null;
//            //시작시각
//            string ToDate = null;
//            //종료일자
//            string ToTime = null;
//            //종료시각
//            //UPGRADE_NOTE: Object이(가) Object_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//            string Object_Renamed = null;
//            //목적
//            string UserSign = null;
//            //UserSign

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//            //사업장
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
//            //사원번호
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            MSTNAM = Strings.Trim(oForm.Items.Item("MSTNAM").Specific.VALUE);
//            //사원성명
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FrDate = Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE);
//            //시작일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FrTime = Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE);
//            //시작시각
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ToDate = Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE);
//            //종료일자
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ToTime = Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE);
//            //종료시각
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Object_Renamed = Strings.Trim(oForm.Items.Item("Object").Specific.VALUE);
//            //목적
//            UserSign = Convert.ToString(MDC_Globals.oCompany.UserSignature);

//            //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
//            sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY032A]";
//            RecordSet01.DoQuery(sQry);

//            if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0)
//            {
//                DocEntry = 1;
//            }
//            else
//            {
//                DocEntry = Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1;
//            }

//            sQry = "                EXEC [PH_PY032_02] ";
//            sQry = sQry + "'" + DocEntry + "',";
//            //관리번호
//            sQry = sQry + "'" + CLTCOD + "',";
//            //사업장
//            sQry = sQry + "'" + MSTCOD + "',";
//            //사원번호
//            sQry = sQry + "'" + MSTNAM + "',";
//            //사원성명
//            sQry = sQry + "'" + FrDate + "',";
//            //시작일자
//            sQry = sQry + "'" + FrTime + "',";
//            //시작시각
//            sQry = sQry + "'" + ToDate + "',";
//            //종료일자
//            sQry = sQry + "'" + ToTime + "',";
//            //종료시각
//            sQry = sQry + "'" + Object_Renamed + "',";
//            //목적
//            sQry = sQry + "'" + UserSign + "'";
//            //UserSign

//            RecordSet02.DoQuery(sQry);

//            MDC_Com.MDC_GF_Message(ref "등록 완료!", ref "S");

//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet02 = null;
//            functionReturnValue = true;
//            return functionReturnValue;
//        PH_PY032_AddData_Error:

//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            functionReturnValue = false;
//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet02 = null;
//            MDC_Com.MDC_GF_Message(ref "PH_PY032_AddData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            return functionReturnValue;
//        }

//        private bool PH_PY032_HeaderSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            //******************************************************************************
//            //Function ID : PH_PY032_HeaderSpaceLineDel()
//            //해당모듈 : PH_PY032
//            //기능 : 필수입력사항 체크
//            //인수 : 없음
//            //반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short ErrNum = 0;
//            ErrNum = 0;

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            switch (true)
//            {
//                case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE)):
//                    //사원번호
//                    ErrNum = 3;
//                    goto PH_PY032_HeaderSpaceLineDel_Error;
//                    break;
//                case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE)):
//                    //시작일자
//                    ErrNum = 4;
//                    goto PH_PY032_HeaderSpaceLineDel_Error;
//                    break;
//                case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE)):
//                    //시작시각
//                    ErrNum = 5;
//                    goto PH_PY032_HeaderSpaceLineDel_Error;
//                    break;
//                case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE)):
//                    //종료일자
//                    ErrNum = 6;
//                    goto PH_PY032_HeaderSpaceLineDel_Error;
//                    break;
//                case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE)):
//                    //종료시각
//                    ErrNum = 7;
//                    goto PH_PY032_HeaderSpaceLineDel_Error;
//                    break;
//            }

//            functionReturnValue = true;
//            return functionReturnValue;
//        PH_PY032_HeaderSpaceLineDel_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            if (ErrNum == 3)
//            {
//                MDC_Com.MDC_GF_Message(ref "사원번호는 필수사항입니다. 확인하세요.", ref "E");
//                oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            else if (ErrNum == 4)
//            {
//                MDC_Com.MDC_GF_Message(ref "시작일자는 필수사항입니다. 확인하세요.", ref "E");
//                oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            else if (ErrNum == 5)
//            {
//                MDC_Com.MDC_GF_Message(ref "시작시각은 필수사항입니다. 확인하세요.", ref "E");
//                oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            else if (ErrNum == 6)
//            {
//                MDC_Com.MDC_GF_Message(ref "종료일자는 필수사항입니다. 확인하세요.", ref "E");
//                oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            else if (ErrNum == 7)
//            {
//                MDC_Com.MDC_GF_Message(ref "종료시각은 필수사항입니다. 확인하세요.", ref "E");
//                oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            functionReturnValue = false;
//            return functionReturnValue;
//        }

//        /// 메트릭스 필수 사항 check
//        private bool PH_PY032_MatrixSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement


//            int i = 0;
//            short ErrNum = 0;
//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            string sQry = null;

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            functionReturnValue = true;
//            return functionReturnValue;
//        PH_PY032_MatrixSpaceLineDel_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            if (ErrNum == 1)
//            {
//                MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//            }
//            else if (ErrNum == 2)
//            {
//                MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", ref "E");
//            }
//            else if (ErrNum == 3)
//            {
//                MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", ref "E");
//            }
//            else if (ErrNum == 4)
//            {
//                MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", ref "E");
//            }
//            else if (ErrNum == 5)
//            {
//                MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", ref "E");
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }
//            functionReturnValue = false;
//            return functionReturnValue;
//        }

//        private void PH_PY032_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            short i = 0;
//            short ErrNum = 0;
//            string sQry = null;

//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            string sCLTCOD = null;

//            switch (oUID)
//            {

//                case "SCLTCOD":

//                    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    sCLTCOD = Strings.Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE);

//                    //UPGRADE_WARNING: oForm.Items(STeamCd).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (oForm.Items.Item("STeamCd").Specific.ValidValues.Count > 0)
//                    {
//                        //UPGRADE_WARNING: oForm.Items(STeamCd).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        for (i = oForm.Items.Item("STeamCd").Specific.ValidValues.Count - 1; i >= 0; i += -1)
//                        {
//                            //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("STeamCd").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                        }
//                    }

//                    //부서콤보세팅
//                    //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("STeamCd").Specific.ValidValues.Add("%", "전체");
//                    sQry = "            SELECT      U_Code AS [Code],";
//                    sQry = sQry + "                 U_CodeNm As [Name]";
//                    sQry = sQry + "  FROM       [@PS_HR200L]";
//                    sQry = sQry + "  WHERE      Code = '1'";
//                    sQry = sQry + "                 AND U_UseYN = 'Y'";
//                    sQry = sQry + "                 AND U_Char2 = '" + sCLTCOD + "'";
//                    sQry = sQry + "  ORDER BY  U_Seq";
//                    MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("STeamCd").Specific), ref sQry, ref "", ref false, ref false);
//                    //UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("STeamCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//                    break;

//            }

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            return;
//        PH_PY032_FlushToItemValue_Error:

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            MDC_Com.MDC_GF_Message(ref "PH_PY032_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");

//        }

//        ///폼의 아이템 사용지정
//        public void PH_PY032_FormItemEnabled()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//            {

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//                MDC_SetMod.CLTCOD_Select(oForm, "SCLTCOD");

//                //        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
//                //        oForm.Items("ItemCode").Enabled = True

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//            {

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//                MDC_SetMod.CLTCOD_Select(oForm, "SCLTCOD");

//                //        oForm.Items("ItemCode").Enabled = True

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//            {

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//                MDC_SetMod.CLTCOD_Select(oForm, "SCLTCOD");

//            }

//            return;
//        PH_PY032_FormItemEnabled_Error:

//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            MDC_Com.MDC_GF_Message(ref "PH_PY032_FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
//        }

//        ///아이템 변경 이벤트
//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                    ////1
//                    Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                    ////5
//                    Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//                    ////7
//                    Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//                    ////8
//                    Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                    ////10
//                    Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//                    ////18
//                    break;
//                ////et_FORM_ACTIVATE
//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//                    ////19
//                    break;
//                ////et_FORM_DEACTIVATE
//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//                    ////20
//                    Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//                    ////27
//                    Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//                    ////4
//                    break;
//                ////et_LOST_FOCUS
//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                    ////17
//                    Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
//                    break;
//            }
//            return;
//        Raise_FormItemEvent_Error:
//            ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            string sQry = null;
//            SAPbobsCOM.Recordset RecordSet01 = null;
//            RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            ////BeforeAction = True
//            if ((pVal.BeforeAction == true))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1284":
//                        //취소
//                        break;
//                    case "1286":
//                        //닫기
//                        break;
//                    case "1293":
//                        //행삭제
//                        break;
//                    case "1281":
//                        //찾기
//                        break;
//                    case "1282":
//                        //추가
//                        ///추가버튼 클릭시 메트릭스 insertrow

//                        PH_PY032_FormReset();

//                        //                oMat01.Clear
//                        //                oMat01.FlushToDataSource
//                        //                oMat01.LoadFromDataSource

//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        BubbleEvent = false;
//                        PH_PY032_LoadCaption();

//                        //oForm.Items("GCode").Click ct_Regular


//                        return;

//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        //레코드이동버튼
//                        break;
//                }
//                ////BeforeAction = False
//            }
//            else if ((pVal.BeforeAction == false))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1284":
//                        //취소
//                        break;
//                    case "1286":
//                        //닫기
//                        break;
//                    case "1293":
//                        //행삭제
//                        break;
//                    case "1281":
//                        //찾기
//                        break;
//                    ////Call PH_PY032_FormItemEnabled '//UDO방식
//                    case "1282":
//                        //추가
//                        break;
//                    //                oMat01.Clear
//                    //                oDS_PH_PY032A.Clear

//                    //                Call PH_PY032_LoadCaption
//                    //                Call PH_PY032_FormItemEnabled
//                    ////Call PH_PY032_FormItemEnabled '//UDO방식
//                    ////Call PH_PY032_AddMatrixRow(0, True) '//UDO방식
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        //레코드이동버튼
//                        break;
//                        ////Call PH_PY032_FormItemEnabled
//                }
//            }
//            return;
//        Raise_FormMenuEvent_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            ////BeforeAction = True
//            if ((BusinessObjectInfo.BeforeAction == true))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//                ////BeforeAction = False
//            }
//            else if ((BusinessObjectInfo.BeforeAction == false))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            return;
//        Raise_FormDataEvent_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//            }
//            if (pVal.ItemUID == "Mat01")
//            {
//                if (pVal.Row > 0)
//                {
//                    oLastItemUID01 = pVal.ItemUID;
//                    oLastColUID01 = pVal.ColUID;
//                    oLastColRow01 = pVal.Row;
//                }
//            }
//            else
//            {
//                oLastItemUID01 = pVal.ItemUID;
//                oLastColUID01 = "";
//                oLastColRow01 = 0;
//            }
//            return;
//        Raise_RightClickEvent_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {

//                if (pVal.ItemUID == "PH_PY032")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                    {
//                    }
//                }

//                ///추가/확인 버튼클릭
//                if (pVal.ItemUID == "BtnAdd")
//                {

//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {

//                        if (PH_PY032_HeaderSpaceLineDel() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }

//                        //                If PH_PY032_DataCheck() = False Then
//                        //                    BubbleEvent = False
//                        //                    Exit Sub
//                        //                End If

//                        if (PH_PY032_AddData() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }

//                        PH_PY032_FormReset();
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                        PH_PY032_LoadCaption();
//                        PH_PY032_MTX01();

//                        oLast_Mode = oForm.Mode;

//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {

//                        if (PH_PY032_HeaderSpaceLineDel() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }

//                        //                If PH_PY032_DataCheck() = False Then
//                        //                    BubbleEvent = False
//                        //                    Exit Sub
//                        //                End If

//                        if (PH_PY032_UpdateData() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }

//                        PH_PY032_FormReset();
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                        PH_PY032_LoadCaption();
//                        PH_PY032_MTX01();

//                        //                oForm.Items("GCode").Click ct_Regular
//                    }

//                    ///조회
//                }
//                else if (pVal.ItemUID == "BtnSearch")
//                {

//                    PH_PY032_FormReset();
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                    ///fm_VIEW_MODE

//                    PH_PY032_LoadCaption();
//                    PH_PY032_MTX01();

//                    ///삭제
//                }
//                else if (pVal.ItemUID == "BtnDelete")
//                {

//                    if (MDC_Globals.Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
//                    {

//                        PH_PY032_DeleteData();
//                        PH_PY032_FormReset();
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        ///fm_VIEW_MODE

//                        PH_PY032_LoadCaption();
//                        PH_PY032_MTX01();

//                    }
//                    else
//                    {

//                    }

//                }
//                else if (pVal.ItemUID == "BtnPrint")
//                {

//                    PH_PY032_Print_Report01();

//                }

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemUID == "PH_PY032")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                    {
//                    }
//                }
//            }

//            return;
//        Raise_EVENT_ITEM_PRESSED_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {

//                MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", "");
//                //기본정보-사번

//                MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SMSTCOD", "");
//                //조회조건-사번

//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }

//            return;
//        Raise_EVENT_KEY_DOWN_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {

//                if (pVal.ItemUID == "Mat01")
//                {

//                    if (pVal.Row > 0)
//                    {

//                        oMat01.SelectRow(pVal.Row, true, false);

//                        oForm.Freeze(true);

//                        //DataSource를 이용하여 각 컨트롤에 값을 출력
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE);
//                        //관리번호
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.VALUE);
//                        //사업장
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE);
//                        //사원번호
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.VALUE);
//                        //사원성명
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_FrDate", 0, Strings.Replace(oMat01.Columns.Item("FrDate").Cells.Item(pVal.Row).Specific.VALUE, ".", ""));
//                        //시작일자
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_FrTime", 0, oMat01.Columns.Item("FrTime").Cells.Item(pVal.Row).Specific.VALUE);
//                        //시작시각
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_ToDate", 0, Strings.Replace(oMat01.Columns.Item("ToDate").Cells.Item(pVal.Row).Specific.VALUE, ".", ""));
//                        //종료일자
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_ToTime", 0, oMat01.Columns.Item("ToTime").Cells.Item(pVal.Row).Specific.VALUE);
//                        //종료시각
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oDS_PH_PY032A.SetValue("U_Object", 0, oMat01.Columns.Item("Object").Cells.Item(pVal.Row).Specific.VALUE);
//                        //목적

//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//                        PH_PY032_LoadCaption();

//                        oForm.Freeze(false);

//                    }
//                }
//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }

//            return;
//        Raise_EVENT_CLICK_Error:

//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {

//                PH_PY032_FlushToItemValue(pVal.ItemUID);

//            }

//            return;
//        Raise_EVENT_COMBO_SELECT_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }
//            return;
//        Raise_EVENT_DOUBLE_CLICK_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }
//            return;
//        Raise_EVENT_MATRIX_LINK_PRESSED_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            if (pVal.BeforeAction == true)
//            {

//                if (pVal.ItemChanged == true)
//                {

//                    if ((pVal.ItemUID == "Mat01"))
//                    {
//                        //                If (pVal.ColUID = "ItemCode") Then
//                        //                    '//기타작업
//                        //                    Call oDS_PH_PY032B.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.VALUE)
//                        //                    If oMat01.RowCount = pVal.Row And Trim(oDS_PH_PY032B.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
//                        //                        PH_PY032_AddMatrixRow (pVal.Row)
//                        //                    End If
//                        //                Else
//                        //                    Call oDS_PH_PY032B.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.VALUE)
//                        //                End If
//                    }
//                    else
//                    {

//                        PH_PY032_FlushToItemValue(pVal.ItemUID);

//                        if (pVal.ItemUID == "MSTCOD")
//                        {

//                            //UPGRADE_WARNING: oForm.Items(MSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'");
//                            //성명

//                        }
//                        else if (pVal.ItemUID == "SMSTCOD")
//                        {

//                            //UPGRADE_WARNING: oForm.Items(SMSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("SMSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.VALUE + "'");
//                            //성명

//                        }

//                    }
//                    //            oMat01.LoadFromDataSource
//                    //            oMat01.AutoResizeColumns
//                    //            oForm.Update
//                }

//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }

//            oForm.Freeze(false);

//            return;
//        Raise_EVENT_VALIDATE_Error:

//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                PH_PY032_FormItemEnabled();
//                ////Call PH_PY032_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
//            }
//            return;
//        Raise_EVENT_MATRIX_LOAD_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                PH_PY032_FormResize();
//            }
//            return;
//        Raise_EVENT_RESIZE_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                //        If (pVal.ItemUID = "ItemCode") Then
//                //            Dim oDataTable01 As SAPbouiCOM.DataTable
//                //            Set oDataTable01 = pVal.SelectedObjects
//                //            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
//                //            Set oDataTable01 = Nothing
//                //        End If
//                //        If (pVal.ItemUID = "CardCode" Or pVal.ItemUID = "CardName") Then
//                //            Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY032A", "U_CardCode,U_CardName")
//                //        End If
//            }
//            return;
//        Raise_EVENT_CHOOSE_FROM_LIST_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.ItemUID == "Mat01")
//            {
//                if (pVal.Row > 0)
//                {
//                    oLastItemUID01 = pVal.ItemUID;
//                    oLastColUID01 = pVal.ColUID;
//                    oLastColRow01 = pVal.Row;
//                }
//            }
//            else
//            {
//                oLastItemUID01 = pVal.ItemUID;
//                oLastColUID01 = "";
//                oLastColRow01 = 0;
//            }
//            return;
//        Raise_EVENT_GOT_FOCUS_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                SubMain.RemoveForms(oFormUniqueID01);
//                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oForm = null;
//                //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oMat01 = null;
//            }
//            return;
//        Raise_EVENT_FORM_UNLOAD_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            int i = 0;
//            if ((oLastColRow01 > 0))
//            {
//                if (pVal.BeforeAction == true)
//                {
//                    //            If (PH_PY032_Validate("행삭제") = False) Then
//                    //                BubbleEvent = False
//                    //                Exit Sub
//                    //            End If
//                    ////행삭제전 행삭제가능여부검사
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    for (i = 1; i <= oMat01.VisualRowCount; i++)
//                    {
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
//                    }
//                    oMat01.FlushToDataSource();
//                    oDS_PH_PY032A.RemoveRecord(oDS_PH_PY032A.Size - 1);
//                    oMat01.LoadFromDataSource();
//                    if (oMat01.RowCount == 0)
//                    {
//                        PH_PY032_Add_MatrixRow(0);
//                    }
//                    else
//                    {
//                        if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY032A.GetValue("U_CntcCode", oMat01.RowCount - 1))))
//                        {
//                            PH_PY032_Add_MatrixRow(oMat01.RowCount);
//                        }
//                    }
//                }
//            }
//            return;
//        Raise_EVENT_ROW_DELETE_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private bool PH_PY032_CreateItems()
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            string oQuery01 = null;
//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oDS_PH_PY032A = oForm.DataSources.DBDataSources("@PH_PY032A");
//            oDS_PH_PY032B = oForm.DataSources.DBDataSources("@PS_USERDS01");

//            //// 메트릭스 개체 할당
//            oMat01 = oForm.Items.Item("Mat01").Specific;
//            oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            oMat01.AutoResizeColumns();

//            //관리번호
//            oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

//            //사업장
//            oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

//            //부서
//            oForm.DataSources.UserDataSources.Add("STeamCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("STeamCd").Specific.DataBind.SetBound(true, "", "STeamCd");

//            //사원번호
//            oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

//            //사원성명
//            oForm.DataSources.UserDataSources.Add("SMSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SMSTNAM").Specific.DataBind.SetBound(true, "", "SMSTNAM");

//            //시작월
//            oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

//            //종료월
//            oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

//            //목적
//            oForm.DataSources.UserDataSources.Add("SObject", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("SObject").Specific.DataBind.SetBound(true, "", "SObject");

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY032_CreateItems_Error:

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        ///콤보박스 set
//        public void PH_PY032_ComboBox_Setting()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            SAPbouiCOM.ComboBox oCombo = null;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet01 = null;

//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oForm.Freeze(true);

//            ////////////조회정보//////////
//            //    '부서
//            //    Call oForm.Items("STeamCd").Specific.ValidValues.Add("%", "전체")
//            //    sQry = "            SELECT      U_Code AS [Code],"
//            //    sQry = sQry & "                 U_CodeNm As [Name]"
//            //    sQry = sQry & "  FROM       [@PS_HR200L]"
//            //    sQry = sQry & "  WHERE      Code = '1'"
//            //    sQry = sQry & "                 AND U_UseYN = 'Y'"
//            //    sQry = sQry & "  ORDER BY  U_Seq"
//            //    Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCd").Specific, sQry, "", False, False)
//            //    Call oForm.Items("STeamCd").Specific.Select(0, psk_Index)

//            ////////////매트릭스//////////
//            //사업장
//            MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId");

//            oForm.Freeze(false);
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            return;
//        PH_PY032_ComboBox_Setting_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY032_CF_ChooseFromList()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            return;
//        PH_PY032_CF_ChooseFromList_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY032_EnableMenus()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            return;
//        PH_PY032_EnableMenus_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY032_SetDocument(string oFromDocEntry01)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            if ((string.IsNullOrEmpty(oFromDocEntry01)))
//            {
//                PH_PY032_FormItemEnabled();
//                ////Call PH_PY032_AddMatrixRow(0, True) '//UDO방식일때
//            }
//            else
//            {
//                //        oForm.Mode = fm_FIND_MODE
//                //        Call PH_PY032_FormItemEnabled
//                //        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
//                //        oForm.Items("1").Click ct_Regular
//            }
//            return;
//        PH_PY032_SetDocument_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY032_FormResize()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            oMat01.AutoResizeColumns();

//            //    oForm.Items("Static15").Left = oForm.Width - oForm.Items("Static07").Left - 49
//            //    oForm.Items("GrpBox02").Left = oForm.Width - oForm.Items("Static07").Left - 50

//            return;
//        PH_PY032_FormResize_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY032_FormReset()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_FormReset()
//            //해당모듈 : PH_PY032
//            //기능 : 화면 초기화
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            string sQry = null;
//            SAPbobsCOM.Recordset RecordSet01 = null;
//            RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oForm.Freeze(true);

//            //관리번호
//            sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]";
//            RecordSet01.DoQuery(sQry);

//            if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0)
//            {
//                oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(1));
//            }
//            else
//            {
//                oDS_PH_PY032A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1));
//            }

//            string User_BPLID = null;
//            User_BPLID = MDC_PS_Common.User_BPLID();

//            ////////////기준정보//////////
//            oDS_PH_PY032A.SetValue("U_CLTCOD", 0, User_BPLID);
//            //사업장
//            oDS_PH_PY032A.SetValue("U_MSTCOD", 0, "");
//            //사원번호
//            oDS_PH_PY032A.SetValue("U_MSTNAM", 0, "");
//            //사원성명
//            oDS_PH_PY032A.SetValue("U_FrDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyyMMdd"));
//            //시작일자
//            oDS_PH_PY032A.SetValue("U_FrTime", 0, "");
//            //시작시각
//            oDS_PH_PY032A.SetValue("U_ToDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyyMMdd"));
//            //종료일자
//            oDS_PH_PY032A.SetValue("U_ToTime", 0, "");
//            //종료시각
//            oDS_PH_PY032A.SetValue("U_Object", 0, "");
//            //목적
//            //출장번호
//            //    Call PH_PY032_GetDestNo

//            ////////////조회정보//////////
//            //    Call oForm.Items("SCLTCOD").Specific.Select(User_BPLID, psk_ByValue) '사업장
//            //    Call oForm.Items("SDestinat").Specific.Select(0, psk_Index) '출장지
//            //    Call oForm.Items("SRegCls").Specific.Select(0, psk_Index) '등록구분
//            //    Call oForm.Items("SObjCls").Specific.Select(0, psk_Index) '목적구분
//            //    Call oForm.Items("SDestCode").Specific.Select(0, psk_Index) '출장지역
//            //    Call oForm.Items("SDestDiv").Specific.Select(0, psk_Index) '출장구분
//            //    Call oForm.Items("SVehicle").Specific.Select(0, psk_Index) '차량구분
//            //    '기간(월)
//            //    oForm.Items("SFrDate").Specific.VALUE = Format(Now, "YYYY.MM")
//            //    oForm.Items("SToDate").Specific.VALUE = Format(Now, "YYYY.MM")

//            oForm.Items.Item("MSTCOD").Click();

//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            oForm.Freeze(false);

//            return;
//        PH_PY032_FormReset_Error:
//            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            RecordSet01 = null;
//            MDC_Com.MDC_GF_Message(ref "PH_PY032_FormReset_Error:" + Err().Number + " - " + Err().Description, ref "E");
//        }

//        private void PH_PY032_CalculateTransExp()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_CalculateTransExp()
//            //해당모듈 : PH_PY032
//            //기능 : 교통비 계산
//            //인수 : 없음
//            //반환값 : 합계 금액
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            short ErrNum = 0;

//            double FuelPrc = 0;
//            //유류단가
//            double Distance = 0;
//            //거리
//            double TransExp = 0;
//            //교통비

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FuelPrc = Convert.ToDouble(Strings.Trim(oForm.Items.Item("FuelPrc").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Distance = Convert.ToDouble(Strings.Trim(oForm.Items.Item("Distance").Specific.VALUE));

//            TransExp = ((FuelPrc * Distance * 0.1) / 10) * 10;
//            //원단위 절사

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("TransExp").Specific.VALUE = TransExp;

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            return;
//        PH_PY032_CalculateTransExp_Error:

//            if (ErrNum == 1)
//            {
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_CalculateTransExp_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//        }

//        private void PH_PY032_CalculateTotalExp()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_CalculateTotalExp()
//            //해당모듈 : PH_PY032
//            //기능 : 합계 금액 계산
//            //인수 : 없음
//            //반환값 : 합계 금액
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            short ErrNum = 0;

//            double TransExp = 0;
//            //교통비
//            double DayExp = 0;
//            //일비
//            double FoodExp = 0;
//            //식비
//            double ParkExp = 0;
//            //주차비
//            double TollExp = 0;
//            //도로비
//            double TotalExp = 0;
//            //합계

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            TransExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("TransExp").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DayExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("DayExp").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FoodExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("FoodExp").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ParkExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("ParkExp").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            TollExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("TollExp").Specific.VALUE));
//            TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp;

//            oDS_PH_PY032A.SetValue("U_TotalExp", 0, Convert.ToString(TotalExp));

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            return;
//        PH_PY032_CalculateTotalExp_Error:

//            if (ErrNum == 1)
//            {
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_CalculateTotalExp_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//        }

//        private void PH_PY032_GetDestNo()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_GetDestNo()
//            //해당모듈 : PH_PY032
//            //기능 : 출장번호 생성
//            //인수 : 없음
//            //반환값 : 합계 금액
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short loopCount = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            short ErrNum = 0;

//            string FrDate = null;
//            string CLTCOD = null;

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FrDate = Strings.Left(Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE), 6);

//            sQry = "EXEC PH_PY032_05 '" + CLTCOD + "', '" + FrDate + "'";
//            oRecordSet01.DoQuery(sQry);

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("DestNo1").Specific.VALUE = FrDate;
//            //UPGRADE_WARNING: oForm.Items(DestNo2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("DestNo2").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("DestNo2").Value);

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            return;
//        PH_PY032_GetDestNo_Error:

//            if (ErrNum == 1)
//            {
//            }
//            else
//            {
//                MDC_Com.MDC_GF_Message(ref "PH_PY032_GetDestNo_Error:" + Err().Number + " - " + Err().Description, ref "E");
//            }

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//        }

//        private void PH_PY032_GetFuelPrc()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_GetFuelPrc()
//            //해당모듈 : PH_PY032
//            //기능 : 유류단가 조회
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short loopCount = 0;
//            string sQry = null;
//            object CheckAmt = null;

//            SAPbobsCOM.Recordset oRecordSet = null;
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            string CLTCOD = null;
//            string StdYear = null;
//            string StdMonth = null;
//            string FuelType = null;

//            double FuelPrice = 0;

//            CLTCOD = Strings.Trim(oDS_PH_PY032A.GetValue("U_CLTCOD", 0));
//            //사업장
//            StdYear = Strings.Left(Strings.Trim(oDS_PH_PY032A.GetValue("U_FrDate", 0)), 4);
//            StdMonth = Strings.Mid(Strings.Trim(oDS_PH_PY032A.GetValue("U_FrDate", 0)), 5, 2);
//            FuelType = Strings.Trim(oDS_PH_PY032A.GetValue("U_FuelType", 0));
//            //유류

//            sQry = "           SELECT      T0.U_Year AS [StdYear],";
//            sQry = sQry + "                T1.U_Month AS [StdMonth],";
//            sQry = sQry + "                T1.U_Gasoline AS [Gasoline],";
//            sQry = sQry + "                T1.U_Diesel AS [Diesel],";
//            sQry = sQry + "                T1.U_LPG AS [LPG]";
//            sQry = sQry + " FROM       [@PH_PY007A] AS T0";
//            sQry = sQry + "                INNER JOIN";
//            sQry = sQry + "                [@PH_PY007B] AS T1";
//            sQry = sQry + "                    ON T0.Code = T1.Code";
//            sQry = sQry + " WHERE      T0.U_CLTCOD = '" + CLTCOD + "'";
//            sQry = sQry + "                AND T0.U_Year = '" + StdYear + "'";
//            sQry = sQry + "                AND T1.U_Month = '" + StdMonth + "'";

//            oRecordSet.DoQuery(sQry);

//            //휘발유
//            if (FuelType == "1")
//            {
//                //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                FuelPrice = oRecordSet.Fields.Item("Gasoline").Value;
//                //가스
//            }
//            else if (FuelType == "2")
//            {
//                //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                FuelPrice = oRecordSet.Fields.Item("LPG").Value;
//                //경유
//            }
//            else if (FuelType == "3")
//            {
//                //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                FuelPrice = oRecordSet.Fields.Item("Diesel").Value;
//            }
//            else
//            {
//                FuelPrice = 0;
//            }

//            oDS_PH_PY032A.SetValue("U_FuelPrc", 0, Convert.ToString(FuelPrice));
//            oForm.Items.Item("FuelPrc").Click();

//            //    oForm.Items("FuelPrc").Specific.VALUE = FuelPrice

//            return;
//        PH_PY032_GetFuelPrc_Error:

//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_GetFuelPrc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY032_CalculateFoodExp()
//        {
//            //******************************************************************************
//            //Function ID : PH_PY032_CalculateFoodExp()
//            //해당모듈 : PH_PY032
//            //기능 : 식비 계산
//            //인수 : 없음
//            //반환값 : 없음
//            //특이사항 : 없음
//            //******************************************************************************
//            // ERROR: Not supported in C#: OnErrorStatement


//            short loopCount = 0;
//            string sQry = null;
//            short ErrNum = 0;

//            SAPbobsCOM.Recordset oRecordSet = null;
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            string MSTCOD = null;
//            //사번
//            short FoodNum = 0;
//            //식수
//            double FoodPrc = 0;
//            //당일식비
//            double FoodExp = 0;
//            //전체식비

//            //사번을 선택하지 않으면
//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY032A.GetValue("U_MSTCOD", 0))) & Strings.Trim(oDS_PH_PY032A.GetValue("U_FoodNum", 0)) != "0")
//            {
//                ErrNum = 1;
//                goto PH_PY032_CalculateFoodExp_Error;
//            }

//            MSTCOD = Strings.Trim(oDS_PH_PY032A.GetValue("U_MSTCOD", 0));
//            //사번
//            FoodNum = Convert.ToInt16(Strings.Trim(oDS_PH_PY032A.GetValue("U_FoodNum", 0)));
//            //식수

//            sQry = "            SELECT      T1.U_Num4 AS [FoodPrc]";
//            sQry = sQry + "  FROM       [@PH_PY001A] AS T0";
//            sQry = sQry + "                 LEFT JOIN";
//            sQry = sQry + "                 [@PS_HR200L] AS T1";
//            sQry = sQry + "                     ON T0.U_JIGCOD = T1.U_Code";
//            sQry = sQry + "                     AND T1.Code = 'P232'";
//            sQry = sQry + "                     AND T1.U_UseYN = 'Y'";
//            sQry = sQry + "  WHERE      T0.Code = '" + MSTCOD + "'";

//            oRecordSet.DoQuery(sQry);

//            //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            FoodPrc = oRecordSet.Fields.Item("FoodPrc").Value;
//            FoodExp = FoodPrc * FoodNum;

//            oDS_PH_PY032A.SetValue("U_FoodExp", 0, Convert.ToString(FoodExp));
//            oForm.Items.Item("FoodExp").Click();

//            return;
//        PH_PY032_CalculateFoodExp_Error:

//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            if (ErrNum == 1)
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("사원을 먼저 선택하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oDS_PH_PY032A.SetValue("U_FoodNum", 0, "0");
//                //0식 선택
//                oForm.Items.Item("MSTCOD").Click();
//            }
//            else
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_CalculateFoodExp_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//        }


//        private void PH_PY032_Print_Report01()
//        {

//            string DocNum = null;
//            short ErrNum = 0;
//            string WinTitle = null;
//            string ReportName = null;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            short DocEntry = 0;
//            string CLTCOD = null;

//            SAPbouiCOM.ProgressBar ProgBar01 = null;
//            ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            /// ODBC 연결 체크
//            if (ConnectODBC() == false)
//            {
//                goto PH_PY032_Print_Report01_Error;
//            }

//            ////인자 MOVE , Trim 시키기..
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = Convert.ToInt16(Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE));
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);

//            /// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//            WinTitle = "[PH_PY032] 사용외출증";

//            //창원
//            if (CLTCOD == "1")
//            {
//                ReportName = "PH_PY032_01.rpt";
//                //동래
//            }
//            else if (CLTCOD == "2")
//            {
//                ReportName = "PH_PY032_02.rpt";
//                //사상
//            }
//            else if (CLTCOD == "3")
//            {
//                ReportName = "PH_PY032_03.rpt";
//            }
//            MDC_Globals.gRpt_Formula = new string[3];
//            MDC_Globals.gRpt_Formula_Value = new string[3];
//            MDC_Globals.gRpt_SRptSqry = new string[2];
//            MDC_Globals.gRpt_SRptName = new string[2];
//            MDC_Globals.gRpt_SFormula = new string[2, 2];
//            MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//            //// Formula 수식필드

//            //// SubReport


//            MDC_Globals.gRpt_SFormula[1, 1] = "";
//            MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

//            /// Procedure 실행"
//            sQry = "EXEC [PH_PY032_90] '" + DocEntry + "'";

//            oRecordSet.DoQuery(sQry);

//            //    If oRecordSet.RecordCount = 0 Then
//            //        ErrNum = 1
//            //        GoTo PH_PY032_Print_Report01_Error
//            //    End If

//            if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false)
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//            ProgBar01.Value = 100;
//            ProgBar01.Stop();
//            //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgBar01 = null;

//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return;
//        PH_PY032_Print_Report01_Error:


//            if (ErrNum == 1)
//            {
//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
//            }
//            else
//            {

//                ProgBar01.Value = 100;
//                ProgBar01.Stop();
//                //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                ProgBar01 = null;

//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY032_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//        }
//    }
//}
