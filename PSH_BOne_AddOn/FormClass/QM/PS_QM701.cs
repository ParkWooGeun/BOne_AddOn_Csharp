using System;
using System.IO;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Scripting;



namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 외주자료등록(부적합)
    /// </summary>
    internal class PS_QM701 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_QM701H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_QM701L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM701.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_QM701_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_QM701");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_QM701_CreateItems();
                PS_QM701_ComboBox_Setting();
                PS_QM701_EnableMenus();
                PS_QM701_SetDocument(oFromDocEntry01);
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
        private void PS_QM701_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_QM701H = oForm.DataSources.DBDataSources.Item("@PS_QM701H");
                oDS_PS_QM701L = oForm.DataSources.DBDataSources.Item("@PS_QM701L");

                oMat01 = oForm.Items.Item("oMat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oDS_PS_QM701H.SetValue("U_WorkDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PS_QM701H.SetValue("U_OrdDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                //내주외주
                oForm.Items.Item("InOut").Specific.ValidValues.Add("I", "자체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("O", "외주");
                //검사자
                oDS_PS_QM701H.SetValue("U_WorkName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("WorkCode").Specific.Value + "'", 0, 1));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_QM701_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("CLTCOD").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //불량코드
                sQry = "SELECT b.U_MidCode, b.U_MidName From [@PS_PP002H] a Inner Join [@PS_PP002L] b On a.DocEntry = b.DocEntry Where a.U_BigCode = '1' Order by b.U_MidCode";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BadCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                // 구매구분
                sQry = " SELECT '%' AS [Code],"; 
                sQry += " '선택' AS [Name]";
                sQry += " UNION ALL";
                sQry += " SELECT Code, ";
                sQry += " Name ";
                sQry += " FROM [@PSH_ORDTYP] ";
                sQry += " ORDER BY Code";

                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("ItemCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("ItemCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 불량원인
                sQry = " SELECT '%' AS [Code],";
                sQry += " '선택' AS [Name]";
                sQry += " UNION ALL";
                sQry += " SELECT U_Code AS [Code],U_CodeNm AS [Name] FROM [@PS_QM700L] WHERE Code ='InCase' ORDER BY Code";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BadNote").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("BadNote").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 판정의견
                sQry = " SELECT '%' AS [Code],";
                sQry += " '선택' AS [Name]";
                sQry += " UNION ALL";
                sQry += " SELECT U_Code AS [Code],U_CodeNm AS [Name] FROM [@PS_QM700L] WHERE Code ='InOpinio' ORDER BY Code";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("verdict").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("verdict").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }
               
        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_QM701_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_QM701_FormItemEnabled();
                    PS_QM701_AddMatrixRow(0, true); //UDO방식일때
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_QM701_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }


        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_QM701_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_QM701_FormClear();
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdDate").Enabled = true;
                    oForm.Items.Item("BadCode").Enabled = true;
                    oForm.Items.Item("InCpCode").Enabled = true;
                    oForm.Items.Item("OuCpCode").Enabled = true;
                    oForm.Items.Item("BadNote").Enabled = true;
                    oForm.Items.Item("verdict").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("cmt").Enabled = true;
                    oForm.EnableMenu("1281", true);  //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("OrdDate").Enabled = false;
                    oForm.Items.Item("BadCode").Enabled = false;
                    oForm.Items.Item("InCpCode").Enabled = false;
                    oForm.Items.Item("OuCpCode").Enabled = false;
                    oForm.Items.Item("BadNote").Enabled = false;
                    oForm.Items.Item("verdict").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;
                    oForm.Items.Item("cmt").Enabled = false;
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true);  //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdDate").Enabled = true;
                    oForm.Items.Item("BadCode").Enabled = true;
                    oForm.Items.Item("InCpCode").Enabled = true;
                    oForm.Items.Item("OuCpCode").Enabled = true;
                    oForm.Items.Item("BadNote").Enabled = true;
                    oForm.Items.Item("verdict").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("cmt").Enabled = true;
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true);  //추가
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
        /// 화면 초기화
        /// </summary>
        private void PS_QM701_FormReset()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                oDS_PS_QM701H.SetValue("U_CLTCOD", 0, dataHelpClass.User_BPLID()); // 사업장
                oDS_PS_QM701H.SetValue("U_WorkDate", 0, DateTime.Now.ToString("yyyyMMdd"));                               
                oDS_PS_QM701H.SetValue("U_KeyDoc", 0, "");                                 
                oDS_PS_QM701H.SetValue("U_WorkNum", 0, "");                             
                oDS_PS_QM701H.SetValue("U_WorkCode", 0, "");                            
                oDS_PS_QM701H.SetValue("U_InOut", 0, "");                              
                oDS_PS_QM701H.SetValue("U_ItemCode",0, "");                             
                oDS_PS_QM701H.SetValue("U_ItemName",0, "");                                
                oDS_PS_QM701H.SetValue("U_ItemSpec", 0, "");
                oDS_PS_QM701H.SetValue("U_CardCode", 0, "");                               
                oDS_PS_QM701H.SetValue("U_InDate", 0, DateTime.Now.ToString("yyyyMMdd"));  
                oDS_PS_QM701H.SetValue("U_TotalQty", 0, "");                               
                oDS_PS_QM701H.SetValue("U_BZZadQty", 0, "");                               
                oDS_PS_QM701H.SetValue("U_OutUnit", 0, "");
                oDS_PS_QM701H.SetValue("U_MSTCOD", 0, "");
                oDS_PS_QM701H.SetValue("U_BadCode", 0, "");                              
                oDS_PS_QM701H.SetValue("U_OrdDate", 0, DateTime.Now.ToString("yyyyMMdd"));                               
                oDS_PS_QM701H.SetValue("U_InCpCode", 0, "");                             
                oDS_PS_QM701H.SetValue("U_OuCpCode", 0, "");                              
                oDS_PS_QM701H.SetValue("U_BadNote", 0, "%");                               
                oDS_PS_QM701H.SetValue("U_verdict", 0, "%");                
                oDS_PS_QM701H.SetValue("U_Comments", 0, "");                              
                oDS_PS_QM701H.SetValue("U_cmt", 0, "");               
                oDS_PS_QM701H.SetValue("U_ChkYN", 0, "");             

                //라인 초기화
                oMat01.Clear();
                oMat01.FlushToDataSource();
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
        /// PS_QM701_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_QM701_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)//행추가여부
                {
                    oDS_PS_QM701L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_QM701L.Offset = oRow;
                oDS_PS_QM701L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_QM701_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM701'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_QM701 사진보여주기
        /// </summary>
        private void PS_QM701_DisplayFixData(string DocEntry)
        {
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (!string.IsNullOrEmpty(DocEntry.ToString().Trim()))
                {
                    sQry = "EXEC PS_QM701_03 '" + DocEntry + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "추가를 먼저 누르신 뒤 사진을 등록해주세요.";
                        throw new Exception();
                    }
                    oDS_PS_QM701H.SetValue("U_Pic", 0, "");
                    oDS_PS_QM701H.SetValue("U_Pic", 0, "\\\\191.1.1.220\\Incom_Pic\\" + oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim() + "_Out.BMP");
                    oDS_PS_QM701H.SetValue("DocEntry", 0, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("Canceled", 0, oRecordSet.Fields.Item("Canceled").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_ChkYN", 0, oRecordSet.Fields.Item("ChkYN").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_CLTCOD", 0, oRecordSet.Fields.Item("CLTCOD").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_KeyDoc", 0, oRecordSet.Fields.Item("KeyDoc").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_WorkDate", 0, oRecordSet.Fields.Item("WorkDate").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_WorkCode", 0, oRecordSet.Fields.Item("WorkCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_WorkName", 0, oRecordSet.Fields.Item("WorkName").Value.ToString().Trim()); 
                    oDS_PS_QM701H.SetValue("U_InOut", 0, oRecordSet.Fields.Item("InOut").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_ItemCode", 0, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_WorkNum", 0, oRecordSet.Fields.Item("WorkNum").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_ItemSpec", 0, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_CardCode", 0, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_CardName", 0, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_InDate", 0, oRecordSet.Fields.Item("InDate").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_TotalQty", 0, oRecordSet.Fields.Item("TotalQty").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_BZZadQty", 0, oRecordSet.Fields.Item("BZZadQty").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_MSTCOD", 0, oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_MSTNAM", 0, oRecordSet.Fields.Item("MSTNAM").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_OrdDate", 0, oRecordSet.Fields.Item("OrdDate").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_BadCode", 0, oRecordSet.Fields.Item("BadCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_InCpCode", 0, oRecordSet.Fields.Item("InCpCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_InCpName", 0, oRecordSet.Fields.Item("InCpName").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_OuCpCode", 0, oRecordSet.Fields.Item("OuCpCode").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_BadNote", 0, oRecordSet.Fields.Item("BadNote").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_verdict", 0, oRecordSet.Fields.Item("verdict").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_Comments", 0, oRecordSet.Fields.Item("Comments").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_OutUnit", 0, oRecordSet.Fields.Item("OutUnit").Value.ToString().Trim());
                    oDS_PS_QM701H.SetValue("U_cmt", 0, oRecordSet.Fields.Item("Cmt").Value.ToString().Trim());
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }


        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_QM701_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (oForm.Items.Item("BadNote").Specific.Value.ToString().Trim() =="%")
                {
                    errMessage = "불량원인을 선택하세요.";
                    throw new Exception();
                }
                if (oForm.Items.Item("verdict").Specific.Value.ToString().Trim() == "%")
                {
                    errMessage = "판정의견을 선택하세요.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("WorkName").Specific.Value))
                {
                    errMessage = "검사자가 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
                {
                    errMessage = "결재자가 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("WorkDate").Specific.Value))
                {
                    errMessage = "검사일자가 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("WorkNum").Specific.Value))
                {
                    errMessage = "작업번호 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                {
                    errMessage = "품목구분 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("BZZadQty").Specific.Value))
                {
                    errMessage = "부적합량 입력되지 않았습니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("OrdDate").Specific.Value))
                {
                    errMessage = "요구납기일 입력되지 않았습니다.";
                    throw new Exception();
                }
                
                if (string.IsNullOrEmpty(oForm.Items.Item("KeyDoc").Specific.Value))
                {
                    errMessage = "검수입고 문서가 선택되지않았습니다.. 다시확인해주세요.";
                    throw new Exception();
                }
                if (float.Parse(oForm.Items.Item("BZZadQty").Specific.Value) > float.Parse(oForm.Items.Item("TotalQty").Specific.Value))
                {
                    errMessage = "부적합량이 입고량보다 많습니다. 확인해주세요.";
                    throw new Exception();
                }
                if (!string.IsNullOrEmpty(oForm.Items.Item("InCpCode").Specific.Value))
                {
                    string sQry = "SELECT CoUNT(*) FROM [@PS_QM700L] WHERE Code = 'OutCode' AND U_Code ='" + (oForm.Items.Item("InCpCode").Specific.Value.ToString().Trim()) + "'";
                    oRecordSet.DoQuery(sQry);
                    if(oRecordSet.Fields.Item(0).Value == 0)
                    {
                        errMessage = "발생공정값을 다시 확인해주세요.";
                        throw new Exception();
                    }
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// OpenFileSelectDialog 호출(쓰레드를 이용하여 비동기화)
        /// OLE 호출을 수행하려면 현재 스레드를 STA(단일 스레드 아파트) 모드로 설정해야 합니다.
        /// </summary>
        [STAThread]
        private string OpenFileSelectDialog()
        {
            string returnFileName = string.Empty;

            var thread = new System.Threading.Thread(() =>
            {
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "bmp Files|*.bmp|All Files|*.*";
                openFileDialog.FilterIndex = 1; //FilterIndex는 1부터 시작
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    returnFileName = openFileDialog.FileName;
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            return returnFileName;
        }

        /// <summary>
        /// PS_QM701_LoadPic
        /// </summary>
        private bool PS_QM701_LoadPic(string pPictureControlName)
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;
            string sFilePath;
            string sFileName;
            string sFileExtension;
            string sQry;
            string SaveFolders;
            FileSystemObject FSO = new FileSystemObject();
            FileListBoxForm fileListBoxForm = new FileListBoxForm();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_QM701H.GetValue("U_WorkNum", 0).ToString().Trim()))
                {
                    throw new Exception();
                }
                SaveFolders = "\\\\191.1.1.220\\Incom_Pic";

                //사진 불러오기
                sFilePath = OpenFileSelectDialog();
                if (string.IsNullOrEmpty(sFilePath))
                {
                    errMessage = "*.BMP 이미지가 선택되지 않았습니다.";
                    throw new Exception();
                }

                //확장자 확인
                sFileName = Path.GetFileName(sFilePath);
                sFileExtension = Path.GetExtension(sFilePath);
                if (sFileExtension.ToUpper() != ".BMP")
                {
                    errMessage = "BMP 확장자만 가능합니다.";
                    throw new Exception();
                }
                
                string imageFileName = "_Out.BMP";

                //서버에 기존 파일 체크
                FileInfo fileInfo = new FileInfo(SaveFolders + "\\" + sFileName);
                if (fileInfo.Exists)
                {
                    FSO.DeleteFile(SaveFolders + "\\" + sFileName);
                }
                FSO.CopyFile(sFilePath, SaveFolders + "\\" + oDS_PS_QM701H.GetValue("DocEntry", 0).ToString().Trim() + imageFileName);

                sQry = " EXEC [PS_QM701_01] '";
                sQry += oDS_PS_QM701H.GetValue("DocEntry", 0).ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.MessageBox("사진이 업로드 되었습니다.");
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oDS_PS_QM701H.SetValue("U_Pic", 0, (SaveFolders + "\\" + sFileName));
                ReturnValue = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return ReturnValue;
        }


        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_QM701_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string filename;
            string DocEntry;
            string Incom_Pic_Path;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            try
            {
                DocEntry= oForm.Items.Item("DocEntry").Specific.Value.Trim();
                filename = "_Out.bmp";
                Incom_Pic_Path = @"\\191.1.1.220\Incom_Pic\";

                if (System.IO.File.Exists(Incom_Pic_Path + DocEntry + filename))
                {
                    if (System.IO.File.Exists(Incom_Pic_Path + "PIC.bmp") == true)
                    {
                        System.IO.File.Delete(Incom_Pic_Path + "PIC.bmp");
                        System.IO.File.Copy(Incom_Pic_Path + DocEntry + filename, Incom_Pic_Path + "PIC.bmp");
                    }
                    else
                    {
                        System.IO.File.Copy(Incom_Pic_Path + DocEntry + filename, Incom_Pic_Path + "PIC.bmp");
                    }
                }
                else
                {
                    System.IO.File.Delete(Incom_Pic_Path + "PIC.bmp");
                    System.IO.File.Copy(Incom_Pic_Path + "NULL.bmp", Incom_Pic_Path + "PIC.bmp");
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();

                WinTitle = "[PS_QM701] 외주 부적합 자재 통보서";
                ReportName = "PS_QM702_01.rpt";

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); //사업장
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry, "PS_QM702_04"));
                
                formHelpClass.OpenCrystalReport(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_QM701_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string errMessage = string.Empty;
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_QM701_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if(oForm.Items.Item("verdict").Specific.Value.Trim() == "2")
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("Comments").Specific.Value))
                                {
                                    errMessage = "특채는 관련근거가 필수입니다.";
                                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            sQry = "insert into PSHDB_IMG.dbo.ZPS_QM701_PIC(BPLId,FixCode) SELECT ";
                            sQry += "'" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "',";
                            sQry += "'" + oForm.Items.Item("DocEntry").Specific.Value.Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_QM701_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if(oForm.Items.Item("ChkYN").Specific.Value.Trim() == "승인" || oForm.Items.Item("Canceled").Specific.Value.Trim() =="Y")
                            {
                                errMessage = "승인되거나 취소된 문서는 수정할수 없습니다.";
                                PSH_Globals.SBO_Application.MessageBox(errMessage);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                sQry = "UPDATE [@PS_QM701H] SET U_ChkYN = NULL WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                            }
                        }
                    }
                    if (pVal.ItemUID == "btn_Print")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_QM701_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                sQry = "SELECT MAX(DocEntry) AS DocEntry FROM [@PS_QM701H]";
                                oRecordSet.DoQuery(sQry);
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                PS_QM701_FormItemEnabled();
                                oForm.Items.Item("DocEntry").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim();
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
                if (pVal.ItemUID == "oMat01")
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "oMat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Pic")
                    {
                        if (PS_QM701_LoadPic(pVal.ItemUID) == true)
                        {
                            //oDS_PS_QM701H.SetValue("U_Pic", 0, "\\\\191.1.1.220\\Incom_Pic\\" +  oForm.Items.Item("DocEntry").Specific.Value + ".BMP");
                            PS_QM701_DisplayFixData(oForm.Items.Item("DocEntry").Specific.Value);
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            try
            {
                if (pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "WorkCode")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("WorkCode").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "KeyDoc")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("KeyDoc").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "InCpCode")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("InCpCode").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "oMat01")
                        {
                            if (pVal.ColUID == "spec")
                            {
                                oMat01.FlushToDataSource();
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_QM701L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_QM701_AddMatrixRow(pVal.Row, false);
                                }
                                oMat01.LoadFromDataSource();
                            }
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        else
                        {
                            if (pVal.ItemUID == "WorkCode")
                            {
                                oForm.Items.Item("WorkName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("WorkCode").Specific.Value + "'", ""); //검사자
                            }
                            if (pVal.ItemUID == "MSTCOD")
                            {
                                oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", ""); //검사자
                            }
                            if (pVal.ItemUID == "InCpCode")
                            {
                                oForm.Items.Item("InCpName").Specific.Value = dataHelpClass.GetValue("SELECT U_CodeNm FROM [@PS_QM700L] WHERE Code = 'OutCode' AND U_Code ='" + oForm.Items.Item("InCpCode").Specific.Value + "'", 0, 1);
                            }
                            if (pVal.ItemUID == "KeyDoc")
                            {
                                sQry = " EXEC PS_QM701_02 '" + oForm.Items.Item("KeyDoc").Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                oDS_PS_QM701H.SetValue("U_WorkNum", 0, oRecordSet01.Fields.Item("WorkNum").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_WorkCode", 0, oRecordSet01.Fields.Item("WorkCode").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_WorkName", 0, oRecordSet01.Fields.Item("WorkName").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_InOut", 0, oRecordSet01.Fields.Item("InOut").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_ItemName", 0, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_ItemCode", 0, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_ItemSpec", 0, oRecordSet01.Fields.Item("ItemSpec").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_CardCode", 0, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_CardName", 0, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_InDate", 0, oRecordSet01.Fields.Item("InDate").Value);
                                oDS_PS_QM701H.SetValue("U_TotalQty", 0, oRecordSet01.Fields.Item("TotalQty").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_BZZadQty", 0, oRecordSet01.Fields.Item("BZZadQty").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_BadCode", 0, oRecordSet01.Fields.Item("BadCode").Value.ToString().Trim());
                                oDS_PS_QM701H.SetValue("U_WorkDate", 0, oRecordSet01.Fields.Item("WorkDate").Value);
                                oDS_PS_QM701H.SetValue("U_OutUnit", 0, oRecordSet01.Fields.Item("OutUnit").Value.ToString().Trim());
                            }
                        }
                    }
                    oForm.Update();
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PS_QM701_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM701H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM701L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                int i = 0;
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_QM701L.RemoveRecord(oDS_PS_QM701L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_QM701_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_QM701L.GetValue("U_spec", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_QM701_AddMatrixRow(oMat01.RowCount, false);
                            }
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_QM701_FormItemEnabled(); //UDO방식
                            break;
                        case "1282": //추가
                            PS_QM701_FormItemEnabled();
                            PS_QM701_FormReset();
                            PS_QM701_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_QM701_FormItemEnabled();
                            break;
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
        }
    }
}
