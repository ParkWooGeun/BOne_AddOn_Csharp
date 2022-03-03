using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전문직 평가
    /// </summary>
    internal class PS_HR416 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Grid oGrid01;
        private SAPbouiCOM.Grid oGrid02;
        private SAPbouiCOM.Grid oGrid03;
        private SAPbouiCOM.DBDataSource oDS_PS_HR416H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_HR416L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocEntry01;

        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR416.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_HR416_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_HR416");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_HR416_CreateItems();
                PS_HR416_ComboBox_Setting();

                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), false); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1284"), false); // 취소
                oForm.EnableMenu(("1293"), false); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_HR416_CreateItems()
        {
            try
            {
                oDS_PS_HR416L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oGrid02 = oForm.Items.Item("Grid02").Specific;
                oGrid03 = oForm.Items.Item("Grid03").Specific;

                oForm.DataSources.DataTables.Add("ZTEMP1");
                oForm.DataSources.DataTables.Add("ZTEMP2");
                oForm.DataSources.DataTables.Add("ZTEMP3");

                oGrid01.DataTable = oForm.DataSources.DataTables.Item("ZTEMP1");
                oGrid02.DataTable = oForm.DataSources.DataTables.Item("ZTEMP2");
                oGrid03.DataTable = oForm.DataSources.DataTables.Item("ZTEMP3");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_HR416_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oForm.Items.Item("Number").Specific.ValidValues.Add("1", "1차평가");
                oForm.Items.Item("Number").Specific.ValidValues.Add("2", "2차평가");
                oForm.Items.Item("Number").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("1", "1차평가자");
                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("2", "2차평가자");
                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("3", "종합평가자");
                oForm.Items.Item("Evaluate").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //사용여부
                oMat01.Columns.Item("Grade").ValidValues.Add("", "선택");
                oMat01.Columns.Item("Grade").ValidValues.Add("S", "S");
                oMat01.Columns.Item("Grade").ValidValues.Add("A", "A");
                oMat01.Columns.Item("Grade").ValidValues.Add("B", "B");
                oMat01.Columns.Item("Grade").ValidValues.Add("C", "C");
                oMat01.Columns.Item("Grade").ValidValues.Add("D", "D");

                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);//아이디별 사업장 세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_HR416_HeaderSpaceLineDel()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    errMessage = "평가년도는 필수사항입니다. 확인하여 주십시오.";
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                {
                    errMessage = "평가자사번은 필수사항입니다. 확인하여 주십시오.";
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("EmpNo").Specific.Value.ToString().Trim()))
                {
                    errMessage = "피평가자가 선택되지 않았습니다. 확인하여 주십시오.";
                }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return ReturnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_HR416_MatrixSpaceLineDel()
        {
            bool ReturnValue = false;
            int i; 
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하여 주십시오.";
                    throw new Exception();
                }
                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (Convert.ToDouble(oDS_PS_HR416L.GetValue("U_ColQty02", i)) == 0)
                        {
                            errMessage = "평가가 다되지 않았습니다.확인하여 주십시오.";
                            throw new Exception();
                        }
                    }
                }
                oMat01.LoadFromDataSource();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return ReturnValue;
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_HR416_FormResize()
        {
            try
            {
                oForm.Items.Item("Grid01").Top = 79;
                oForm.Items.Item("Grid01").Height = (oForm.Height / 2) - 110;
                oForm.Items.Item("Grid01").Left = 10;
                oForm.Items.Item("Grid01").Width = oForm.Width / 2;

                oForm.Items.Item("Mat01").Top = 79;
                oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 110;
                oForm.Items.Item("Mat01").Left = (oForm.Width / 2) + 30;
                oForm.Items.Item("Mat01").Width = (oForm.Width / 2) - 30;


                oForm.Items.Item("Btn02").Top = oForm.Items.Item("Grid01").Height + 79;
                oForm.Items.Item("2").Top = oForm.Items.Item("Grid01").Height + 79;
                oForm.Items.Item("Btn03").Top = oForm.Items.Item("Grid01").Height + 79;
                oForm.Items.Item("Btn04").Top = oForm.Items.Item("Grid01").Height + 79;
                oForm.Items.Item("Btn05").Top = oForm.Items.Item("Grid01").Height + oForm.Items.Item("Btn04").Height + 79;

                oForm.Items.Item("t3").Top = oForm.Items.Item("Btn02").Top;
                oForm.Items.Item("t3").Left = oForm.Items.Item("Btn01").Left;
                oForm.Items.Item("Svalue").Top = oForm.Items.Item("Btn02").Top;
                oForm.Items.Item("Svalue").Left = oForm.Items.Item("Btn01").Left + 80;

                oForm.Items.Item("t2").Top = 61;
                oForm.Items.Item("t2").Left = (oForm.Width / 2) + 30;

                oForm.Items.Item("t4").Top = oForm.Items.Item("Btn02").Top + 20;

                oForm.Items.Item("Grid02").Top = oForm.Items.Item("Grid01").Height + 120;
                oForm.Items.Item("Grid02").Height = 126;
                oForm.Items.Item("Grid02").Left = 10;
                oForm.Items.Item("Grid02").Width = oForm.Width - 21;

                oForm.Items.Item("t5").Top = oForm.Items.Item("Grid02").Top + 126;

                oForm.Items.Item("Grid03").Top = oForm.Items.Item("t5").Top + 20;
                oForm.Items.Item("Grid03").Height = 126;
                oForm.Items.Item("Grid03").Left = 10;
                oForm.Items.Item("Grid03").Width = oForm.Width - 21;

                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Search_Grid_Data
        /// </summary>
        private void PS_HR416_Search_Grid_Data()
        {
            int Cnt;
            string errMessage = string.Empty;
            string sQry = null;
            string BPLID;
            string Year_Renamed;
            string Number;
            string Evaluate;
            string MSTCOD;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oMat01.Clear();
                oForm.Items.Item("Svalue").Specific.Value = 0;
                oForm.Items.Item("Code").Specific.Value = "";
                oForm.Items.Item("EmpNo").Specific.Value = "";
                oForm.Items.Item("EmpName").Specific.Value = "";
                oForm.Items.Item("CName").Specific.Value = "";

                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                Evaluate = oForm.Items.Item("Evaluate").Specific.Value;
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                if (Evaluate == "1")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD1,'') = '" + MSTCOD + "'";
                    sQry += " and Isnull(b.U_Complet1,'N') = 'Y' ";
                }
                else if (Evaluate == "2")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD2,'') = '" + MSTCOD + "'";
                    sQry += " and Isnull(b.U_Complet2,'N') = 'Y' ";
                }
                else if (Evaluate == "3")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD3,'') = '" + MSTCOD + "'";
                    sQry += " and Isnull(b.U_Complet3,'N') = 'Y' ";
                }

                oRecordSet01.DoQuery(sQry);
                Cnt = oRecordSet01.Fields.Item(0).Value;

                if (Cnt > 0)
                {
                    oForm.Items.Item("Complete").Specific.Value = "평가완료처리";
                }
                else
                {
                    oForm.Items.Item("Complete").Specific.Value = "평가완료미처리";
                }

                if (Evaluate == "1")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD1,'') = '" + MSTCOD + "'";
                }
                else if (Evaluate == "2")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD2,'') = '" + MSTCOD + "'";
                    sQry += " and Isnull(b.U_Complet1,'N') = 'N' ";
                }
                else if (Evaluate == "3")
                {
                    sQry = " select COUNT(*) ";
                    sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD3,'') = '" + MSTCOD + "'";
                    sQry += " and Isnull(b.U_Complet2,'N') = 'N' ";
                }

                oRecordSet01.DoQuery(sQry);
                Cnt = oRecordSet01.Fields.Item(0).Value;

                if (Evaluate == "1")
                {
                    if (Cnt <= 0)
                    {
                        errMessage = "평가대상자가 없습니다.";
                        throw new Exception();
                    }
                }
                else if (Evaluate == "2")
                {
                    if (Cnt > 0)
                    {
                        errMessage = "1차평가가 완료되지 않았습니다.";
                        throw new Exception();
                    }
                }

                sQry = "EXEC PS_HR416_01 '" + BPLID + "','" + Year_Renamed + "','" + Number + "','" + Evaluate + "','" + MSTCOD + "'";
                oGrid01.DataTable.ExecuteQuery(sQry);
                PS_HR416_GridSetting();
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_HR416_MTX01
        /// </summary>
        private void PS_HR416_Search_Grid_Data1(int oRow)
        {
            string errMessage = string.Empty;
            string sQry;
            string RateCode;
            string BPLID;
            string MSTCOD;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                RateCode = oMat01.Columns.Item("RateCode").Cells.Item(oRow).Specific.Value;
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("EmpNo").Specific.Value; //평가대상자

                sQry = "EXEC PS_HR416_03 '" + BPLID + "','" + RateCode + "'";

                oGrid02.DataTable.ExecuteQuery(sQry);
                PS_HR416_GridSetting();

                sQry = "EXEC PS_HR416_04 '" + BPLID + "', '" + MSTCOD + "', '" + RateCode + "'";

                oGrid03.DataTable.ExecuteQuery(sQry);
                PS_HR416_GridSetting();
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// //
        /// </summary>
        /// <param name="A_B"></param>
        /// <param name="Cancel"></param>
        /// <returns></returns>
        private bool PS_HR416_Save_Data(string A_B, bool Cancel)
        {
            bool returnValue = false;
            int i;
            int Cnt;
            int Evaluate;
            string sQry;
            string BPLID;
            string Year_Renamed;
            string Number;
            string Code;
            string EmpNo;
            double Svalue = 0;
            double Avg;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                Code = oForm.Items.Item("Code").Specific.Value;
                EmpNo = oForm.Items.Item("EmpNo").Specific.Value;
                Evaluate = oForm.Items.Item("Evaluate").Specific.Value;

                sQry = "Select Count(*) From Z_PS_HR410L Where Code = '" + Code + "' And MSTCOD = '" + EmpNo + "' and Evaluate = '" + Evaluate + "'";
                oRecordSet01.DoQuery(sQry);

                Cnt = oRecordSet01.Fields.Item(0).Value;

                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        oDS_PS_HR416L.Offset = i;
                        if (Cnt <= 0)
                        {
                            sQry = "insert into Z_PS_HR410L values (";
                            sQry += "'" + Code + "','" + EmpNo + "','" + Evaluate + "',";
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_LineNum", i).ToString().Trim() + "',"; //LineNum
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColReg01", i).ToString().Trim() + "',"; //RateCode
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColReg02", i).ToString().Trim() + "',";
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColReg03", i).ToString().Trim() + "',";
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColQty01", i).ToString().Trim() + "',";
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColReg04", i).ToString().Trim() + "',";
                            sQry += "'" + oDS_PS_HR416L.GetValue("U_ColQty02", i).ToString().Trim() + "')";
                            oRecordSet01.DoQuery(sQry);

                            if (oDS_PS_HR416L.GetValue("U_ColReg01", i).ToString().Trim() != "A12")
                            {
                                Svalue += Convert.ToDouble(oDS_PS_HR416L.GetValue("U_ColQty02", i).ToString().Trim());
                            }
                        }
                        else
                        {
                            sQry = "Update Z_PS_HR410L";
                            sQry += " Set Grade = '" + oDS_PS_HR416L.GetValue("U_ColReg04", i).ToString().Trim() + "' ,";
                            sQry += " Value = '" + oDS_PS_HR416L.GetValue("U_ColQty02", i).ToString().Trim() + "'";
                            sQry += " Where Code = '" + Code + "' And MSTCOD = '" + EmpNo + "' And Evaluate = '" + Evaluate + "'";
                            sQry += " And LineNum = '" + oDS_PS_HR416L.GetValue("U_LineNum", i).ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            if (oDS_PS_HR416L.GetValue("U_ColReg01", i).ToString().Trim() != "A12")
                            {
                                Svalue += Convert.ToDouble(oDS_PS_HR416L.GetValue("U_ColQty02", i).ToString().Trim());
                            }
                        }
                    }

                    switch (Evaluate)
                    {
                        case 1://1차
                            Avg = System.Math.Round(Svalue * 3 / 10, 1);

                            sQry = "Update [@PS_HR410L] ";
                            sQry += " Set U_Value1 = '" + Svalue + "', U_Avg1 = '" + Avg + "'";
                            sQry += " From [@PS_HR410H] a ";
                            sQry += " Where a.Code = [@PS_HR410L].Code ";
                            sQry += " And a.U_BPLId = '" + BPLID + "' and a.U_Year = '" + Year_Renamed + "'";
                            sQry += " And a.U_Number = '" + Number + "' and [@PS_HR410L].U_MSTCOD = '" + EmpNo + "'";
                            oRecordSet01.DoQuery(sQry);
                            break;

                        case 2: //2차
                            Avg = System.Math.Round(Svalue * 4 / 10, 1);

                            sQry = "Update [@PS_HR410L] ";
                            sQry += " Set U_Value2 = '" + Svalue + "', U_Avg2 = U_Avg1 + '" + Avg + "'";
                            sQry += " From [@PS_HR410H] a ";
                            sQry += " Where a.Code = [@PS_HR410L].Code ";
                            sQry += " And a.U_BPLId = '" + BPLID + "' and a.U_Year = '" + Year_Renamed + "'";
                            sQry += " And a.U_Number = '" + Number + "' and [@PS_HR410L].U_MSTCOD = '" + EmpNo + "'";
                            oRecordSet01.DoQuery(sQry);
                            break;
                        case 3 : //3차
                            Avg = System.Math.Round(Svalue * 3 / 10, 1);

                            sQry = "Update [@PS_HR410L] ";
                            sQry += " Set U_Value3 = '" + Svalue + "', U_Avg3 = U_Avg2 + '" + Avg + "'";
                            sQry += " From [@PS_HR410H] a ";
                            sQry += " Where a.Code = [@PS_HR410L].Code ";
                            sQry += " And a.U_BPLId = '" + BPLID + "' and a.U_Year = '" + Year_Renamed + "'";
                            sQry += " And a.U_Number = '" + Number + "' and [@PS_HR410L].U_MSTCOD = '" + EmpNo + "'";
                            oRecordSet01.DoQuery(sQry);
                            break;
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// GridSetting
        /// </summary>
        /// <returns></returns>
        private bool PS_HR416_Complete_Data()
        {
            bool ReturnValue = false;
            int Cnt;
            int Evaluate;
            string sQry = string.Empty;
            string BPLID;
            string Year_Renamed;
            string Number;
            string Code;
            string MSTCOD;
            string Complete;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                Code = oForm.Items.Item("Code").Specific.Value;
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                Evaluate = oForm.Items.Item("Evaluate").Specific.Value;
                Complete = oForm.Items.Item("Complete").Specific.Value;

                if (Evaluate == 1)
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" + Evaluate + "'";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD1,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(c.Value,0) = 0 ";
                }
                else if (Evaluate == 2)
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" + Evaluate + "'";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD2,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(c.Value,0) = 0 ";
                }
                else if (Evaluate == 3)
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" + Evaluate + "'";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD3,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(c.Value,0) = 0 ";
                }

                oRecordSet01.DoQuery(sQry);
                Cnt = oRecordSet01.Fields.Item(0).Value;

                if (Cnt > 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("평가자의 평가가 완료되지 않았습니다.확인바랍니다.");
                }
                else
                {
                    if (Evaluate == 1)
                    {
                        //평가가 전부 다되었으면
                        sQry = " Update [@PS_HR410L] set U_Complet1 = 'Y' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD1 = '" + MSTCOD + "'";
                    }
                    else if (Evaluate == 2)
                    {
                        sQry = " Update [@PS_HR410L] set U_Complet2 = 'Y' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD2 = '" + MSTCOD + "'";
                    }
                    else if (Evaluate == 3)
                    {
                        sQry = " Update [@PS_HR410L] set U_Complet3 = 'Y' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD3 = '" + MSTCOD + "'";
                    }
                    oRecordSet01.DoQuery(sQry);
                    PSH_Globals.SBO_Application.MessageBox(Evaluate + "차 평가완료처리했습니다.");
                }
                oForm.Items.Item("Btn01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
            return ReturnValue;
        }

        /// <summary>
        /// GridSetting
        /// </summary>
        /// <returns></returns>
        private bool PS_HR416_InComplete_Data()
        {
            bool ReturnValue = false;
            string sQry = null;
            int Cnt;
            int Evaluate = 0;
            string BPLID;
            string Year_Renamed;
            string Number;
            string Code;
            string MSTCOD;
            string Complete;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                Code = oForm.Items.Item("Code").Specific.Value;
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                Evaluate = oForm.Items.Item("Evaluate").Specific.Value;
                Complete = oForm.Items.Item("Complete").Specific.Value;

                if (Complete == "평가완료미처리")
                {
                    errMessage = "평가완료처리시만 작업할 수 있습니다. 확인바랍니다.";
                    throw new Exception();
                }

                if (Evaluate == 1) //1차계수조정
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD1,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(b.U_AValue1,0) <> 0 ";
                   
                }
                else if (Evaluate == 2) //2차계수조정
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD2,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(b.U_AValue2,0) <> 0 ";
                }
                else if (Evaluate == 3)//3차계수조정
                {
                    sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code";
                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                    sQry += " and a.U_Year ='" + Year_Renamed + "'";
                    sQry += " and a.U_Number = '" + Number + "'";
                    sQry += " and Isnull(b.U_MSTCOD3,'') = '" + MSTCOD + "'";
                    sQry += " and isnull(b.U_AValue3,0) <> 0 ";
                }

                oRecordSet01.DoQuery(sQry);
                Cnt = oRecordSet01.Fields.Item(0).Value;

                if (Cnt > 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("계수조정처리가 되었습니다. 계수조정취소작업을 먼저하시기바랍니다.");
                }
                else
                {
                    if (Evaluate == 1)
                    {
                        sQry = " Update [@PS_HR410L] set U_Complet1 = 'N' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD1 = '" + MSTCOD + "'";
                    }
                    else if (Evaluate == 2)
                    {
                        sQry = " Update [@PS_HR410L] set U_Complet2 = 'N' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD2 = '" + MSTCOD + "'";
                    }
                    else if (Evaluate == 3)
                    {
                        sQry = " Update [@PS_HR410L] set U_Complet3 = 'N' ";
                        sQry += " From [@PS_HR410H] a";
                        sQry += " Where a.Code = [@PS_HR410L].Code";
                        sQry += " And a.U_BPLId = '" + BPLID + "'";
                        sQry += " And a.U_Year = '" + Year_Renamed + "'";
                        sQry += " And a.U_Number = '" + Number + "'";
                        sQry += " And [@PS_HR410L].U_MSTCOD3 = '" + MSTCOD + "'";
                    }
                    oRecordSet01.DoQuery(sQry);
                    PSH_Globals.SBO_Application.MessageBox(Evaluate + "차 평가완료처리 취소했습니다.");
                }
                oForm.Items.Item("Btn01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
            return ReturnValue;
        }

        /// <summary>
        /// GridSetting
        /// </summary>
        /// <returns></returns>
        private void PS_HR416_GridSetting()
        {
            int i;
            string sColsTitle;

            try
            {
                oForm.Freeze(true);
                oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                for (i = 0; i <= oGrid01.Columns.Count - 1; i++)
                {
                    sColsTitle = oGrid01.Columns.Item(i).TitleObject.Caption;

                    oGrid01.Columns.Item(i).Editable = false;

                    if (sColsTitle == "1차" || sColsTitle == "2차" || sColsTitle == "3차" || sColsTitle == "평균")
                    {
                        oGrid01.Columns.Item(i).RightJustified = true;
                    }

                    if (oGrid01.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
                    {
                        oGrid01.Columns.Item(i).RightJustified = true;
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
        /// PasswordChk
        /// </summary>
        /// <returns></returns>
        private bool PS_HR416_PasswordChk(SAPbouiCOM.ItemEvent pVal)
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
        /// PS_HR416_MTX01
        /// </summary>
        private void PS_HR416_Search_Matrix_Data()
        {
            string errMessage = string.Empty;
            int i;
            int j;
            int Cnt;
            string sQry;
            string Code;
            string BPLID;
            string Year_Renamed;
            string MSTCOD;
            string Number;
            string Evaluate;
            double Svalue = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                Evaluate = oForm.Items.Item("Evaluate").Specific.Value;

                oGrid02.DataTable.Clear();
                oGrid03.DataTable.Clear();

                for (i = 0; i <= oGrid01.Rows.Count - 1; i++)
                {
                    if (oGrid01.Rows.IsSelected(i) == true)
                    {
                        Code = oGrid01.DataTable.GetValue(0, i).ToString().Trim();
                        MSTCOD = oGrid01.DataTable.GetValue(5, i).ToString().Trim();
                        oForm.Items.Item("Code").Specific.Value = Code;
                        oForm.Items.Item("EmpNo").Specific.Value = MSTCOD;
                        oForm.Items.Item("EmpName").Specific.Value = oGrid01.DataTable.GetValue(3, i).ToString().Trim();
                        oForm.Items.Item("CName").Specific.Value = oGrid01.DataTable.GetValue(4, i).ToString().Trim();

                        sQry = "EXEC PS_HR416_02 '" + BPLID + "', '" + Year_Renamed + "', '" + MSTCOD + "', '" + Number + "', '" + Evaluate + "'";
                        oRecordSet01.DoQuery(sQry);

                        Cnt = oDS_PS_HR416L.Size;
                        if (Cnt > 0)
                        {
                            for (j = 0; j <= Cnt - 1; j++)
                            {
                                oDS_PS_HR416L.RemoveRecord(oDS_PS_HR416L.Size - 1);
                            }
                            if (Cnt == 1)
                            {
                                oDS_PS_HR416L.Clear();
                            }
                        }
                        oMat01.LoadFromDataSource();
                        
                        j = 1;
                        while (!(oRecordSet01.EoF))
                        {

                            if (oDS_PS_HR416L.Size < j)
                            {
                                oDS_PS_HR416L.InsertRecord(j - 1);
                                //라인추가
                            }
                            oDS_PS_HR416L.SetValue("U_LineNum", j - 1, Convert.ToString(j));
                            oDS_PS_HR416L.SetValue("U_ColReg01", j - 1, oRecordSet01.Fields.Item(0).Value);
                            oDS_PS_HR416L.SetValue("U_ColReg02", j - 1, oRecordSet01.Fields.Item(1).Value);
                            oDS_PS_HR416L.SetValue("U_ColReg03", j - 1, oRecordSet01.Fields.Item(2).Value);
                            oDS_PS_HR416L.SetValue("U_ColQty01", j - 1, oRecordSet01.Fields.Item(3).Value);
                            oDS_PS_HR416L.SetValue("U_ColReg04", j - 1, oRecordSet01.Fields.Item(4).Value);
                            oDS_PS_HR416L.SetValue("U_ColQty02", j - 1, oRecordSet01.Fields.Item(5).Value);
                            
                            Svalue += oRecordSet01.Fields.Item(5).Value;
                            j += 1;
                            oRecordSet01.MoveNext();
                        }
                        oMat01.LoadFromDataSource();
                    }
                }
                if (Svalue != 0)
                {
                    oForm.Items.Item("Btn02").Specific.Caption = "수정";
                }
                else
                {
                    oForm.Items.Item("Btn02").Specific.Caption = "추가";
                }
                oForm.Items.Item("Svalue").Specific.Value = Svalue;
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_HR416_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string BPLID;
            string Year_Renamed;
            string Number;
            string MSTCOD;
            string FULLNAME;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "MSTCOD":
                        BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                        Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                        Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                        MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                        sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.String.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("FULLNAME").Specific.String = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                        FULLNAME = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        sQry = "Select Count(*) From Z_PS_HR403 Where BPLId = '" + BPLID + "' and Year = '" + Year_Renamed + "' And Number = '" + Number + "' and MSTCOD = '" + MSTCOD + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item(0).Value <= 0)
                        {
                            PS_HR403 tempForm = new PS_HR403();
                            tempForm.LoadForm(BPLID, Year_Renamed, Number, MSTCOD, FULLNAME);
                        }
                        break;
                }
                if (oUID == "Mat01")
                {
                    switch (oCol)
                    {
                        case "GRADE":
                            oMat01.FlushToDataSource();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            string Year_Renamed;
            string BPLID;
            string Number;
            string MSTCOD;
            string FULLNAME;
            string Complete;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                        Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                        Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
                        MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                        FULLNAME = oForm.Items.Item("FULLNAME").Specific.Value;

                        //평가서약을 했는지...
                        sQry = "Select Count(*) From Z_PS_HR403 Where BPLId = '" + BPLID + "' and Year = '" + Year_Renamed + "' And Number = '" + Number + "' and MSTCOD = '" + MSTCOD + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item(0).Value <= 0)
                        {
                            //평가서약 화면
                            //ChildForm01.LoadForm(BPLID, Year_Renamed, Number, MSTCOD, FULLNAME);
                        }
                        else
                        {
                            if (PS_HR416_PasswordChk(pVal) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                                oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else
                            {
                                PS_HR416_Search_Grid_Data();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        Complete = oForm.Items.Item("Complete").Specific.Value;

                        if (Complete == "평가완료처리")
                        {
                            PSH_Globals.SBO_Application.MessageBox("이미 평가완료처리를 하였습니다. 확인바랍니다.");
                            PS_HR416_Search_Grid_Data(); //새로고침
                        }
                        else
                        {
                            if (PS_HR416_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_HR416_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //이동요청중량 보다 기포장중량+포장중량이 클수없음, Check!!!
                            if (PS_HR416_Save_Data("A", false) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                PS_HR416_Search_Grid_Data();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Btn04")
                    {
                        Complete = oForm.Items.Item("Complete").Specific.Value;

                        if (Complete == "평가완료처리")
                        {
                            PSH_Globals.SBO_Application.MessageBox("이미 평가완료처리를 하였습니다. 확인바랍니다.");
                        }
                        else
                        {
                            PS_HR416_Complete_Data();
                        }
                    }
                    else if (pVal.ItemUID == "Btn05")
                    {
                        Complete = oForm.Items.Item("Complete").Specific.Value;

                        if (Complete == "평가완료미처리")
                        {
                            PSH_Globals.SBO_Application.MessageBox("평가완료 취소할 자료가 없습니다. 확인바랍니다.");
                        }
                        else
                        {
                            PS_HR416_InComplete_Data();
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
                        if (pVal.ItemUID == "Year")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value))
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
                else if (pVal.Before_Action == false)
                {
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string sQry;
            string Grade;
            string Year_Renamed;
            string RateCode;
            double Svalue = 0;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "Grade")
                        {
                            Year_Renamed = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                            RateCode = oMat01.Columns.Item("RateCode").Cells.Item(pVal.Row).Specific.Value;
                            Grade = oMat01.Columns.Item("Grade").Cells.Item(pVal.Row).Specific.Value;

                            sQry = "SELECT s = b.U_SValue, a = b.U_AValue, b= b.U_BValue, c = U_CValue, d = U_DValue From [@PS_HR400H] a Inner Join [@PS_HR400L] b On a.Code = b.Code ";
                            sQry += "Where a.U_Year = '" + Year_Renamed + "' And b.U_RateCode = '" + RateCode + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.FlushToDataSource();

                            switch (Grade)
                            {
                                case "S":
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                                    break;                                             
                                case "A":                                              
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                                    break;                                             
                                case "B":                                              
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                                    break;                                             
                                case "C":                                              
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, oRecordSet01.Fields.Item(3).Value.ToString().Trim());
                                    break;                                             
                                case "D":                                              
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, oRecordSet01.Fields.Item(4).Value.ToString().Trim());
                                    break;
                                default:
                                    oDS_PS_HR416L.SetValue("U_ColQty02", pVal.Row - 1, Convert.ToString(0));
                                    break;
                            }

                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oDS_PS_HR416L.Offset = i;

                                Svalue += Convert.ToDouble(oDS_PS_HR416L.GetValue("U_ColQty02", i).ToString().Trim());

                            }

                            oForm.Items.Item("Svalue").Specific.Value = Svalue;
                            
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
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
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        PS_HR416_Search_Matrix_Data();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Year")
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                            {
                                sQry = "select U_Number from [@PS_HR410H] a";
                                sQry += " Where Isnull(a.U_OpenYN,'N') = 'Y' and isnull(a.U_CloseYN,'N') = 'N' ";
                                sQry += " and a.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "' ";
                                sQry += " and a.U_Year = '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "' ";
                                oRecordSet01.DoQuery(sQry);

                                if (!string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                                {
                                    oForm.Items.Item("Number").Specific.Select(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                            }
                        }
                        if (pVal.ItemUID == "MSTCOD")
                        {
                            PS_HR416_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "Grade"))
                        {
                            PS_HR416_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_HR416L);
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
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
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
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_HR416_Search_Grid_Data1(pVal.Row);
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    PS_HR416_FormResize();
                }
                else if (pVal.Before_Action == false)
                {
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
                            if (PS_HR416_HeaderSpaceLineDel() == false) 
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_HR416_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if(Convert.ToString(PSH_Globals.SBO_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소")) == "2") 
                            {
                                BubbleEvent = false;
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
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
    }
}
