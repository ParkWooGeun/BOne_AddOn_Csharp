using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 퇴직금 계산
    /// </summary>
    internal class PH_PY115 : PSH_BaseClass
    {
        private string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;

        private SAPbouiCOM.DBDataSource oDS_PH_PY115A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY115B;
        private SAPbouiCOM.Matrix oMat1;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        private bool MsterChk;

        private string FIXTYP;
        private string ROUNDT;
        private short RODLEN;
        private string RETCHK;
        private string RATUSE;
        private string GNSGBN;
        private string RETYCH;
        //private string oOLDCHK;

        private string oCode;

        private string[] WK_SILSIL = new string[9];
        private string[] WK_ROUNDT = new string[9];
        private short[] WK_LENGTH = new short[9];
        private string[] oCLTOPT = new string[3];

        private double[] SHWJIG = new double[4];
        private double[] SHWIYO = new double[4];
        private double[] SHWGSU = new double[4];
        private double[] SHWTIL = new double[2];
        private double[] SHWSUR = new double[2];
        private double[] SHWGON = new double[2];
        private double[] SHWGWA = new double[2];
        private double[] SHWYAG = new double[2];
        private double[] SHWYAS = new double[2];
        private double[] SHWSAN = new double[2];

        private string BubJung_YN;
        private double WT_SHWTOT; //세액환산명세총계

        #region 내부클래스 선언(구조체를 마이그레이션)
        internal class WG01CODR
        {
            internal double[] TB1AMT = new double[100];
            internal double[] TB1GON = new double[100];
            internal double[] TB1RAT = new double[100];
            internal double[] TB1KUM = new double[100];
        }
        WG01CODR WG01 = new WG01CODR(); //클래스 전체에서 사용됨, 클래스 레벨로 인스턴스 생성

        internal class ZPY401_H
        {
            internal double SHWJIG;
            internal double SHWTIL;
            internal double SHWSUR;
            internal double SHWGON;
            internal double SHWGWA;
            internal double SHWYAG;
            internal double SHWYAS;
        }
        ZPY401_H HS = new ZPY401_H();

        internal class ZPY401_T
        {
            internal double SPCGON;
            internal double SODGON;
            internal double RETGON;
            internal double TAXSTD;
            internal double YTXSTD;
            internal double YSANTX;
            internal double SANTAX;
            internal double RTXGON;
            internal double FRNGON;
            internal double GULGAB;
            internal double WSGGON; // 환산급여공제
        }
        ZPY401_T WK401 = new ZPY401_T();
        #endregion

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY115.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY115_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY115");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                // oForm.DataBrowser.BrowseBy = "DocNum"

                oForm.Freeze(true);
                PH_PY115_CreateItems();
                PH_PY115_EnableMenus();
                PH_PY115_SetDocument(oFormDocEntry01);
                PSH_Globals.ExecuteEventFilter(typeof(PH_PY115));
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY115_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY115A = oForm.DataSources.DBDataSources.Item("@PH_PY115A"); //헤더
                oDS_PH_PY115B = oForm.DataSources.DBDataSources.Item("@PH_PY115B"); //라인(급여내역)

                oMat1 = oForm.Items.Item("Mat1").Specific; //@PH_PY115B
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                //oCheck = oForm.Items.Item("Check1").Specific; //잠금
                oForm.Items.Item("Check1").Specific.ValOff = "N";
                oForm.Items.Item("Check1").Specific.ValOn = "Y";

                //oCheck = oForm.Items.Item("SPCGBN").Specific; //법정퇴직주현근무지세액공제여부
                oForm.Items.Item("SPCGBN").Specific.ValOff = "N";
                oForm.Items.Item("SPCGBN").Specific.ValOn = "Y";

                //oCheck = oForm.Items.Item("JS1SPC").Specific; //법정퇴직종전근무지퇴직세액공제여부
                oForm.Items.Item("JS1SPC").Specific.ValOff = "N";
                oForm.Items.Item("JS1SPC").Specific.ValOn = "Y";

                //oCheck = oForm.Items.Item("BUBCHK").Specific; //서식구분
                oForm.Items.Item("BUBCHK").Specific.ValOff = "N";
                oForm.Items.Item("BUBCHK").Specific.ValOn = "Y";

                //oCombo = oForm.Items.Item("JSNGBN").Specific;
                oForm.Items.Item("JSNGBN").Specific.ValidValues.Add("1", "퇴직정산");
                oForm.Items.Item("JSNGBN").Specific.ValidValues.Add("2", "중도정산");

                //oCombo = oForm.Items.Item("RETRES").Specific;
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("1", "정년퇴직");
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("2", "정리해고");
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("3", "자발적퇴직");
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("4", "임원퇴직");
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("5", "중도정산");
                oForm.Items.Item("RETRES").Specific.ValidValues.Add("6", "기    타");

                oForm.Items.Item("Folder1").AffectsFormMode = false;
                //oFolder = oForm.Items.Item("Folder1").Specific;
                oForm.Items.Item("Folder1").Specific.Select();
                oForm.Items.Item("Folder2").AffectsFormMode = false;
                oForm.Items.Item("Folder3").AffectsFormMode = false;
                oForm.Items.Item("Folder4").AffectsFormMode = false;
                oForm.Items.Item("Folder5").AffectsFormMode = false;
                oForm.Items.Item("Folder6").AffectsFormMode = false;
                oForm.Items.Item("Folder1").Enabled = true;
                oForm.Items.Item("Folder2").Enabled = true;
                oForm.Items.Item("Folder3").Enabled = true;
                oForm.Items.Item("Folder4").Enabled = true;
                oForm.Items.Item("Folder5").Enabled = true;
                oForm.Items.Item("Folder6").Enabled = true;

                oForm.PaneLevel = 1;

                //1.1. 공식가져오기-수당
                BaseSetting();

                oMat1.Columns.Item("CSUCOD").Editable = false;
                oMat1.Columns.Item("CSUNAM").Editable = false;
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
        /// 메뉴 세팅(Enable)
        /// </summary>
        private void PH_PY115_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 초기화(Set)
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY115_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (oFormDocEntry01 == "")
                {
                    PH_PY115_FormItemEnabled();
                    PH_PY115_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY115_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 공식 설정
        /// </summary>
        private void BaseSetting()
        {
            string sQry;
            int i;
            short errNum = 0;
            string JSNYER;

            FIXTYP = "0";
            ROUNDT = "0";
            RODLEN = 0;
            RETCHK = "0";

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //2009.10.26 중도정산할경우 2009년에 2005년까지 정산해달라고 하는경우등등 지급하기로 한 약정일의 귀속년도기준으로 봄.
                //해서 귀속기간년도로 변경함.
                JSNYER = codeHelpClass.Left(oDS_PH_PY115A.GetValue("U_ENDINT", 0), 4); //귀속기간종료
                if (JSNYER.Trim() == "")
                {
                    JSNYER = DateTime.Now.ToString("yyyy");
                }

                //91 기초세액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                for (i = 0; i < 100; i++)
                {
                    WG01.TB1AMT[i] = 0;
                    WG01.TB1GON[i] = 0;
                    WG01.TB1RAT[i] = 0;
                    WG01.TB1KUM[i] = 0;
                }

                sQry = "  SELECT      U_CODNBR,";
                sQry += "             U_CODAMT,";
                sQry += "             U_CODGON,";
                sQry += "             U_CODRAT,";
                sQry += "             U_CODKUM ";
                sQry += " FROM        [@PH_PY100B]";
                sQry += " WHERE       Code =  (";
                sQry += "                         SELECT      Top 1";
                sQry += "                                     Code";
                sQry += "                         FROM        [@PH_PY100A]";
                sQry += "                         WHERE       Code <= '" + JSNYER.Trim() + "'";
                sQry += "                         ORDER BY    Code Desc";
                sQry += "                     )";
                sQry += " ORDER BY    Code,";
                sQry += "             Convert(Integer, U_CODNBR) DESC";
                oRecordSet.DoQuery(sQry);

                while (!oRecordSet.EoF)
                {
                    WG01.TB1AMT[Convert.ToInt16(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODAMT").Value);
                    WG01.TB1GON[Convert.ToInt16(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODGON").Value);
                    WG01.TB1RAT[Convert.ToInt16(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODRAT").Value);
                    WG01.TB1KUM[Convert.ToInt16(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODKUM").Value);
                    oRecordSet.MoveNext();
                }

                //타이틀
                sQry = " SELECT TOP 1 * FROM [@PH_PY114A] T1 ORDER BY T1.Code DESC";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    oCode = oRecordSet.Fields.Item("Code").Value;
                    FIXTYP = oRecordSet.Fields.Item("U_FIXTYP").Value; //평균임금산정방법
                    ROUNDT = oRecordSet.Fields.Item("U_ROUNDT").Value; //일할끝전처리
                    RODLEN = Convert.ToInt16(oRecordSet.Fields.Item("U_LENGTH").Value); //일할단위
                    RETCHK = oRecordSet.Fields.Item("U_RETCHK").Value; //비정기상여포함여부
                    RATUSE = oRecordSet.Fields.Item("U_RATUSE").Value; //임원누진율
                    GNSGBN = oRecordSet.Fields.Item("U_GNSGBN").Value; //근속년수기준
                    RETYCH = oRecordSet.Fields.Item("U_RETYCH").Value; //퇴사월연차수당처리
                }

                //라인정보
                i = 0;
                sQry = "  SELECT      T0.U_CSUCOD,";
                sQry += "             T0.U_SILCOD,";
                sQry += "             T0.U_ROUNDT,";
                sQry += "             T0.U_LENGTH";
                sQry += " FROM        [@PH_PY114B] T0";
                sQry += " WHERE       T0.Code = '" + oCode + "'";
                sQry += " ORDER BY    T0.Code,";
                sQry += "             T0.U_CSUCOD";
                oRecordSet.DoQuery(sQry);

                while (!oRecordSet.EoF)
                {
                    WK_SILSIL[i] = oRecordSet.Fields.Item("U_SILCOD").Value; //코딩용계산식
                    WK_ROUNDT[i] = oRecordSet.Fields.Item("U_ROUNDT").Value; //끝전처리
                    WK_LENGTH[i] = Convert.ToInt16(oRecordSet.Fields.Item("U_LENGTH").Value); //단위
                    i += 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("퇴직금기준설정에 대한 데이터가 없습니다. 입력하여주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// 화면(Form) 아이템 세팅(Enable)
        /// </summary>
        private void PH_PY115_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PH_PY115_FormClear();
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// Matirx 행 추가
        /// </summary>
        private void PH_PY115_AddMatrixRow()
        {
            try
            {
                oForm.Freeze(true);

                // '//[Mat1 용]
                // oMat1.FlushToDataSource
                // oRow = oMat1.VisualRowCount
                // 
                // If oMat1.VisualRowCount > 0 Then
                // If Trim(oDS_PH_PY115B.GetValue("U_FILD01", oRow - 1)) <> "" Then
                // If oDS_PH_PY115B.Size <= oMat1.VisualRowCount Then
                // oDS_PH_PY115B.InsertRecord (oRow)
                // End If
                // oDS_PH_PY115B.Offset = oRow
                // oDS_PH_PY115B.setValue "U_LineNum", oRow, oRow + 1
                // oDS_PH_PY115B.setValue "U_FILD01", oRow, ""
                // oDS_PH_PY115B.setValue "U_FILD02", oRow, ""
                // oDS_PH_PY115B.setValue "U_FILD03", oRow, 0
                // oMat1.LoadFromDataSource
                // Else
                // oDS_PH_PY115B.Offset = oRow - 1
                // oDS_PH_PY115B.setValue "U_LineNum", oRow - 1, oRow
                // oDS_PH_PY115B.setValue "U_FILD01", oRow - 1, ""
                // oDS_PH_PY115B.setValue "U_FILD02", oRow - 1, ""
                // oDS_PH_PY115B.setValue "U_FILD03", oRow - 1, 0
                // oMat1.LoadFromDataSource
                // End If
                // ElseIf oMat1.VisualRowCount = 0 Then
                // oDS_PH_PY115B.Offset = oRow
                // oDS_PH_PY115B.setValue "U_LineNum", oRow, oRow + 1
                // oDS_PH_PY115B.setValue "U_FILD01", oRow, ""
                // oDS_PH_PY115B.setValue "U_FILD02", oRow, ""
                // oDS_PH_PY115B.setValue "U_FILD03", oRow, 0
                // oMat1.LoadFromDataSource
                // End If
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
        /// 화면 클리어
        /// </summary>
        private void PH_PY115_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY115'", "");

                if (System.Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("Code").Specific.Value = "00000001";
                }
                else
                {
                    oForm.Items.Item("Code").Specific.Value = DocEntry.PadLeft(8, '0');
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <param name="oUID"></param>
        /// <returns></returns>
        private bool PH_PY115_DataValidCheck(string oUID)
        {
            bool returnValue = false;
            short errNum = 0;
            double TOTAMT;

            try
            {
                //헤더부분 체크_S
                if (oDS_PH_PY115A.GetValue("U_Check", 0).Trim() == "Y") //&& oOLDCHK.Trim() == "Y")
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim() == "")
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim() == "")
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim() == "")
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim()) < Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()))
                {
                    errNum = 5;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_JIGBIL", 0).Trim() == "")
                {
                    errNum = 6;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_SINYMM", 0).Trim() == "")
                {
                    errNum = 7;
                    throw new Exception();
                }
                else if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0)) < Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TAXGON", 0)))
                {
                    errNum = 8;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_STRINT", 0).Trim() == "" || oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim() == "")
                {
                    errNum = 9;
                    throw new Exception();
                }
                else if (oDS_PH_PY115A.GetValue("U_RETRES", 0).Trim() == "")
                {
                    errNum = 10;
                    throw new Exception();
                }
                else if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim()) >= 20091231 && (oDS_PH_PY115A.GetValue("U_SPCGBN", 0).Trim() == "Y" || oDS_PH_PY115A.GetValue("U_JS1SPC", 0).Trim() == "Y"))
                {
                    errNum = 11;
                    throw new Exception();
                }
                //헤더부분 체크_E

                TOTAMT = 0;
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()); //퇴직급여
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim()); //명예퇴직수당(50%)
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BHMAMT", 0).Trim()); //단체퇴직보험
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_YILPA1", 0).Trim()); //법정퇴직연금일시금
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_YILPA2", 0).Trim()); //법정외퇴직연금일시금
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JTOT01", 0).Trim()); //종전퇴직계1
                TOTAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JTOT02", 0).Trim()); //종전퇴직계2

                //RETPAY - 총퇴직금, SH1JIG - 환산-지급예상액(현)
                if (oUID != "Btn2")
                {
                    if (TOTAMT != Convert.ToDouble(oDS_PH_PY115A.GetValue("U_RETPAY", 0).Trim()) && Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1JIG", 0).Trim()) == 0)
                    {
                        errNum = 12;
                        throw new Exception();
                    }
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("정산시작일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("정산종료일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("정산종료일자가 정산시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("신고연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 8)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("세액(외국납부)공제는 산출세액을 초과할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 9)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속기간은 필수입니다. 입력하시기 바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 10)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("퇴직사유는 필수입니다. 선택하시기 바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 11)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("세액공제는 2009년에 한해서만 가능 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 12)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("퇴직금이 변경되었습니다. 세액계산이 필요합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 세액 계산
        /// </summary>
        /// <returns></returns>
        private bool Compute_TAX()
        {
            short errNum = 0;

            //VB6.FixedLengthString STRDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString JINDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString JOTDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString INPDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString ST2RET = new VB6.FixedLengthString(8);

            string STRDAT;
            string JINDAT;
            string JOTDAT;
            string INPDAT;
            string ST2RET;

            string WT_RTXGBN_H;
            string WT_RTXGBN_J;
            int WT_RTXYER_1;
            int WT_RTXYER_2;

            //근속년수
            int WT_GNMYER;
            int WT_JONYER;
            int WT_JMMMON;
            int WT_JMMYER;
            int WT_JDOYER = 0;

            //퇴직급여
            double WT_BHMAMT;
            double WT_TJKPAY;
            double WT_SUDAMT;
            double WT_YILPA1;
            double WT_YILPA2;
            double WT_RETPAY;
            double WT_RETPA1;
            double WT_RETPA2;
            double WT_JRET01;
            double WT_JSUD01;
            //double WT_JYIL01;
            double WT_JYIL03;
            double WT_JYIL04;
            double WT_JTOT01;
            double WT_JTOT02;
            double WT_BTXAMT;

            //세액계산
            //double WT_SPCGON;
            //double WT_SODGON;
            double WT_RETGON;
            double WT_TAXSTD;
            double WT_YTXSTD;
            double WT_YSANTX;
            //double WT_SANTAX;
            double WT_RTXGON;
            double WT_FRNGON;
            double WT_FRNGO1;
            double WT_FRNGO2;
            double WT_GULGAB;
            double WT_GULJUM;

            //세액환산정보
            double[] D_SHWJIG = new double[4]; //세액환산지급예상액
            double[] D_SHWYAS = new double[4]; //세액환산-환산연평균산출세액
            double[] D_SHWSUR = new double[4]; //세액환산-수령액
            double[] D_RTXGON = new double[4]; //퇴직소득세액공제
            double[] D_RTXKUM = new double[4]; //퇴직소득세액공제한도
            double[] D_RTXSAN = new double[4]; //산출세액
            double[] D_GULGAB = new double[2];
            short i;

            bool returnValue = false;

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //Clear
                for (i = 0; i < 4; i++)
                {
                    D_SHWJIG[i] = 0;
                    D_SHWYAS[i] = 0;
                    D_SHWSUR[i] = 0;

                    D_RTXGON[i] = 0;
                    D_RTXKUM[i] = 0;
                    D_RTXSAN[i] = 0;
                }

                //귀속연도
                STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim(); //주현근무지 정산시작일
                INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0).Trim(); //주현근무지 입사일자
                JINDAT = oDS_PH_PY115A.GetValue("U_JINDAT", 0).Trim(); //종전근무지 입사일자
                JOTDAT = oDS_PH_PY115A.GetValue("U_JOTDAT", 0).Trim(); //종전근무지 퇴사일자
                ST2RET = oDS_PH_PY115A.GetValue("U_ST2RET", 0).Trim(); //종전근무지 중도정산일
                //퇴직세액공제여부
                WT_RTXGBN_H = oDS_PH_PY115A.GetValue("U_SPCGBN", 0).Trim(); //2009년주현퇴직소득세액공제여부(Y/N)
                WT_RTXGBN_J = oDS_PH_PY115A.GetValue("U_JS1SPC", 0).Trim(); //2009년 종전퇴직소득세액공제여부(Y/N)

                //************************************************************************************************************
                // 1.근속년월일 계산
                //************************************************************************************************************
                Compute_GNMYER();

                //세법상 근무년수 주(현):18년
                WT_GNMYER = (oDS_PH_PY115A.GetValue("U_GNMMON", 0).Trim() == "" ? 0 : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMMON", 0).Trim())) - (oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim() == "" ? 0 : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim()));

                if ((WT_GNMYER % 12) == 0)
                {
                    WT_GNMYER /= 12;
                }
                else
                {
                    WT_GNMYER = Convert.ToInt16(dataHelpClass.IInt(WT_GNMYER / 12, 1)) + 1;
                }

                //세법상 근무년수 종(전):2년
                WT_JONYER = (oDS_PH_PY115A.GetValue("U_GNMDAY", 0).Trim() == "" ? 0 : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMDAY", 0).Trim())) - (oDS_PH_PY115A.GetValue("U_JEXMON", 0).Trim() == "" ? 0 : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JEXMON", 0).Trim()));

                if ((WT_JONYER % 12) == 0)
                {
                    WT_JONYER /= 12;
                }
                else
                {
                    WT_JONYER = Convert.ToInt16(dataHelpClass.IInt(WT_JONYER / 12, 1)) + 1;
                }

                //세법상 중복월수:25년
                WT_JMMMON = (oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim() == "" ? 0 : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim())); //세법상 중복월수
                WT_JMMYER = WT_JMMMON / 12; //세법상 중복년수

                if ((WT_JMMMON % 12) == 0)
                {
                    WT_JMMYER = WT_JMMMON / 12;
                }
                else
                {
                    WT_JMMYER = Convert.ToInt16(dataHelpClass.IInt(WT_JMMMON / 12, 1)) + 1;
                }

                //중도정산 년도
                if (INPDAT != STRDAT)
                {
                    dataHelpClass.Term2(INPDAT, STRDAT);
                    WT_JDOYER = PSH_Globals.ZPAY_GBL_GNSYER * 12 + PSH_Globals.ZPAY_GBL_GNSMON;

                    if (PSH_Globals.ZPAY_GBL_GNSDAY > 0)
                    {
                        WT_JDOYER += 1;
                    }

                    if ((WT_JDOYER % 12) == 0)
                    {
                        WT_JDOYER /= 12;
                    }
                    else
                    {
                        WT_JDOYER = Convert.ToInt16(dataHelpClass.IInt(WT_JDOYER / 12, 1)) + 1;
                    }
                }

                oDS_PH_PY115A.SetValue("U_GN3YER", 0, WT_JDOYER.ToString());

                //종전근무지에 종전입사일자와 주현의 입사일자가 동일할경우 법정외근속월수도 주현과 동일하게 변경함.
                if (codeHelpClass.Left(JOTDAT, 4) == codeHelpClass.Left(STRDAT, 4))
                {
                    oDS_PH_PY115A.SetValue("U_GN2MON", 0, oDS_PH_PY115A.GetValue("U_GNMMON", 0).Trim());
                }

                //세액 및 계산식 다시 불러오기(정산종료연도를 바꾸거나 네비게이션사용시 폼로드년도 다를수있슴)
                BaseSetting();

                //************************************************************************************************************
                // 2. 퇴직금 지급내역
                //************************************************************************************************************
                //현근무지 소득(법정)
                WT_TJKPAY = (oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim())); //주현(법정) 퇴직급여
                WT_BHMAMT = (oDS_PH_PY115A.GetValue("U_BHMAMT", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BHMAMT", 0).Trim())); //주현(법정) 단체보험료
                WT_YILPA1 = (oDS_PH_PY115A.GetValue("U_YILPA1", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_YILPA1", 0).Trim())); //주현(법정) 퇴직연금일시금

                //현근무지 소득(법정외)
                WT_SUDAMT = (oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim())); //주현(법정외) 명예퇴직수당
                WT_YILPA2 = (oDS_PH_PY115A.GetValue("U_YILPA2", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_YILPA2", 0).Trim())); //주현(법정외) 퇴직연금일시금

                //전근무지 소득(법정)
                WT_JRET01 = (oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim())); //종전(법정) 퇴직급여
                WT_JYIL03 = (oDS_PH_PY115A.GetValue("U_JYIL03", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYIL03", 0).Trim())); //종전(법정) 퇴직연금일시금
                WT_JTOT01 = (oDS_PH_PY115A.GetValue("U_JTOT01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JTOT01", 0).Trim())); //종전(법정) 계

                //전근무지 소득 (법정외)
                WT_JSUD01 = (oDS_PH_PY115A.GetValue("U_JADD01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JADD01", 0).Trim())); //종전(법정외) 명예퇴직수당
                WT_JYIL04 = (oDS_PH_PY115A.GetValue("U_JYIL04", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYIL04", 0).Trim())); //종전(법정외) 퇴직연금일시금
                WT_JTOT02 = (oDS_PH_PY115A.GetValue("U_JTOT02", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JTOT02", 0).Trim())); //종전(법정외) 계

                WT_RETPA1 = WT_TJKPAY + WT_BHMAMT + WT_YILPA1 + WT_JTOT01; //법정퇴직금 계
                WT_RETPA2 = WT_SUDAMT + WT_YILPA2 + WT_JTOT02; //법정외퇴직금계
                WT_RETPAY = WT_RETPA1 + WT_RETPA2; //퇴직급여액[퇴직급여+명예퇴직수당+단체퇴직수당+종전퇴직금](21) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

                //비과세계 (주현+종전)
                WT_BTXAMT = (oDS_PH_PY115A.GetValue("U_BTXPAY", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BTXPAY", 0).Trim())) + (oDS_PH_PY115A.GetValue("U_BTXP01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BTXP01", 0).Trim()));

                //************************************************************************************************************
                // 2.1 체크 옵션
                //************************************************************************************************************
                //명예퇴직금처리여부(조건: 퇴직정산년도이전 년도에 중도정산한 사원의 퇴직정산시 법정외퇴직금이 존재할 경우
                //베루(중도정산임에도 불구하고, 명예퇴직수당을 지급)
                if (Convert.ToInt16(codeHelpClass.Left(INPDAT, 4)) < Convert.ToInt16(codeHelpClass.Left(STRDAT, 4)) && (WT_SUDAMT + WT_JSUD01) > 0)
                {
                    BubJung_YN = "Y";
                }
                else
                {
                    BubJung_YN = "N";
                }

                //주현근무지의 입사일자와 정산시작일이 다를경우 : 중도정산을 한경우
                if ((INPDAT != STRDAT && codeHelpClass.Left(JOTDAT, 4) == codeHelpClass.Left(STRDAT, 4)) || (ST2RET == "" && JOTDAT != ""))
                {
                    BubJung_YN = "N";
                }

                oDS_PH_PY115A.SetValue("U_BUBCHK", 0, BubJung_YN);
                oForm.Items.Item("BUBCHK").Update();

                if ((WT_RETPA1 + WT_SUDAMT + WT_BTXAMT) == 0)
                {
                    errNum = 1;
                    throw new Exception();
                    //PSH_Globals.SBO_Application.StatusBar.SetText("주(현) 퇴직급여가 0 입니다. 퇴직금 계산을 먼저 하시기 바랍니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    //return;
                }

                //************************************************************************************************************
                // 3. 퇴직연금내역
                //퇴직연금에 가입한 근로자가 퇴직연금일시금을 지급받는 경우에 한하여 입력함.
                //************************************************************************************************************

                // 3.1. 퇴직연금명세-퇴직연금일시금 계산(주현, 종전)
                Display_ilSiGum();

                //************************************************************************************************************
                // 4.세액환산명세
                //퇴직연금을 일시금과 연금으로 나누어 수령하는 경우
                //퇴직연금을 수령하던 자가 중도해지등으로 일시금으로 지급받는 경우
                //퇴직소득을 과세이연한 경우 등에 한하여 입력함.
                //************************************************************************************************************

                //세액환산명세 입력받은 값
                WT_SHWTOT = 0;

                for (i = 0; i < 4; i++)
                {
                    SHWJIG[i] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH" + (i + 1) + "JIG", 0).Trim()); //퇴직연금일시금지급예상액
                    SHWIYO[i] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH" + (i + 1) + "IYO", 0).Trim()); //과세이연금액
                    SHWGSU[i] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH" + (i + 1) + "GSU", 0).Trim()); //기수령한퇴직급여액
                    WT_SHWTOT = WT_SHWTOT + SHWJIG[i] + SHWIYO[i] + SHWGSU[i];
                }
                WT_FRNGO1 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1TAX", 0)); //외국납부세액(법정)
                WT_FRNGO2 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2TAX", 0)); //외국납부세액(법정외)
                WT_FRNGON = WT_FRNGO1 + WT_FRNGO2;

                //세액환산명세가 있을경우
                if (WT_SHWTOT > 0)
                {
                    //2.1 세액환산명세-주(현)
                    Display_TaxWhanSan(3, BubJung_YN);
                    D_SHWJIG[2] = HS.SHWJIG;
                    D_SHWYAS[2] = HS.SHWYAS;
                    D_SHWSUR[2] = HS.SHWSUR;

                    //2.2 세액환산명세-종(전)
                    Display_TaxWhanSan(4, BubJung_YN);
                    D_SHWJIG[3] = HS.SHWJIG;
                    D_SHWYAS[3] = HS.SHWYAS;
                    D_SHWSUR[3] = HS.SHWSUR;

                    //2.3 세액환산명세-법정퇴직
                    Display_TaxWhanSan(1, BubJung_YN);
                    D_SHWJIG[0] = HS.SHWJIG;
                    D_SHWYAS[0] = HS.SHWYAS;
                    D_SHWSUR[0] = HS.SHWSUR;

                    //2.4 세액환산명세-법정퇴직이외
                    Display_TaxWhanSan(2, BubJung_YN);
                    D_SHWJIG[1] = HS.SHWJIG;
                    D_SHWYAS[1] = HS.SHWYAS;
                    D_SHWSUR[1] = HS.SHWSUR;

                    // ***************************************************************************
                    // 결과값 UPDATE
                    // ***************************************************************************

                    oDS_PH_PY115A.SetValue("U_SH2GON", 0, WK401.RETGON.ToString()); //퇴직소득공제
                    oDS_PH_PY115A.SetValue("U_SH2GWA", 0, WK401.TAXSTD.ToString()); //퇴직소득과세표준
                    oDS_PH_PY115A.SetValue("U_SH2YAG", 0, WK401.YTXSTD.ToString()); //연평균과세표준
                    oDS_PH_PY115A.SetValue("U_SH2YAS", 0, WK401.YSANTX.ToString()); //연평균산출세액
                    oDS_PH_PY115A.SetValue("U_JS2SAN", 0, WK401.SANTAX.ToString()); //산출세액
                    oDS_PH_PY115A.SetValue("U_JS2GON", 0, WK401.RTXGON.ToString()); //퇴직소득세액공제
                }

                //************************************************************************************************************
                //퇴직소득세액  계산
                //************************************************************************************************************
                WT_RTXYER_1 = (WT_GNMYER < WT_JONYER ? WT_JONYER : WT_GNMYER); //근무년수(법정)
                WT_RTXYER_2 = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GN2YER", 0).Trim()); //근무년수(법정외)

                //주현근무지의 입사일자와 정산시작일이 다를경우 : 중도정산을 한경우
                if (BubJung_YN == "N")
                {
                    WT_RETPA1 += WT_RETPA2; //동일년도중도정산은 종전으로 처리함. 모두 법정퇴직금처리함.
                    WT_RETPA2 = 0;
                }

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //법정퇴직금
                //종전이나 주현에서 중도정산을 하지 않을 경우는 모두 법정퇴직금
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Compute_TAX_2(WT_RETPA1, WT_RTXYER_1, 0, D_SHWJIG[0], D_SHWYAS[0], D_SHWSUR[0], "BB", 0);

                oDS_PH_PY115A.SetValue("U_JS1RET", 0, WT_RETPA1.ToString()); //퇴직급여총액
                oDS_PH_PY115A.SetValue("U_JS1GON", 0, WK401.RETGON.ToString()); //퇴직소득공제
                oDS_PH_PY115A.SetValue("U_JS1GYEAR", 0, WK401.SODGON.ToString()); //퇴직소득근속년수공제 //2013/12/25 노근용추가
                //oDS_PH_PY115A.SetValue("U_JS1GRATE", 0, WK401.SPCGON.ToString()); //퇴직소득정율공제 //2013/12/25 노근용추가 // 20200222 삭제조치(황영수)
                //oDS_PH_PY115A.SetValue("U_JS1STD", 0, WK401.TAXSTD.ToString()); //퇴직소득과세표준 // 20200222 삭제조치(황영수)

                double FD_JS3STD = 0;

                if (Convert.ToInt16(codeHelpClass.Left(oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim(), 4)) >= 2016)
                {
                    //2016년 이후 계산
                    oDS_PH_PY115A.SetValue("U_JS2GYEAR", 0, WK401.SODGON.ToString()); //퇴직소득근속년수공제 //2016/1/6 노근용추가
                    oDS_PH_PY115A.SetValue("U_JS2AMT", 0, ((WT_RETPA1 - WK401.SODGON) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNMYER", 0).Trim()) * 12).ToString()); //환산급여 //2016/1/6 노근용추가

                    oDS_PH_PY115A.SetValue("U_JS2GAMT", 0, WK401.WSGGON.ToString()); //2016년 환산급여별 공제
                    oDS_PH_PY115A.SetValue("U_JS3STD", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2AMT", 0).Trim()) - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2GAMT", 0).Trim())).ToString()); //2016년 퇴직소득과세표준

                    FD_JS3STD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS3STD", 0).Trim()); //2016년 이후 환산산출세액

                    for (i = 51; i <= 59; i++)
                    {
                        if (FD_JS3STD <= WG01.TB1AMT[i])
                        {
                            FD_JS3STD = dataHelpClass.IInt((FD_JS3STD - WG01.TB1GON[i]) * (WG01.TB1RAT[i] / 100) + WG01.TB1KUM[i], 1);
                            break;
                        }
                    }

                    oDS_PH_PY115A.SetValue("U_JS2HYST3", 0, FD_JS3STD.ToString()); //환산산출세액(2016/01/01 이후)
                    oDS_PH_PY115A.SetValue("U_JS2SAN03", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2HYST3", 0).Trim()) / 12 * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNMYER", 0).Trim())).ToString()); //산출세액(2016/01/01 이후)
                    //oDS_PH_PY115A.SetValue("U_TAXYEAR", 0, codeHelpClass.Left((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim())).ToString(), 4)); //퇴직과세연도(2016년 이후)
                }

                //퇴직소득과세표준 안분 (2012/12/31이전, 이후) //2013/12/25 노근용추가
                //oDS_PH_PY115A.SetValue("U_JS1STD01", 0, dataHelpClass.IInt(WK401.TAXSTD * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim()) / (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER2", 0).Trim())), 1).ToString());
                //oDS_PH_PY115A.SetValue("U_JS1STD02", 0, (WK401.TAXSTD - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1STD01", 0).Trim())).ToString());

                //oDS_PH_PY115A.SetValue("U_JS1YSD", 0, WK401.YTXSTD.ToString()); //연평균과세표준

                ////연평균과세표준 안분 (2012/12/31이전, 이후)  //2013/12/25 노근용추가
                //if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim()) > 0)
                //{
                //    oDS_PH_PY115A.SetValue("U_JS1YSD01", 0, dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1STD01", 0).Trim()) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim()), 1).ToString());
                //}
                //else
                //{
                //    oDS_PH_PY115A.SetValue("U_JS1YSD01", 0, "0");
                //}
                //oDS_PH_PY115A.SetValue("U_JS1YSD02", 0, dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1STD02", 0).Trim()) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER2", 0).Trim()), 1).ToString());

                ////환산과세표준(2013/01/01 이후)
                //oDS_PH_PY115A.SetValue("U_JS1HYTD2", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YSD02", 0).Trim()) * 5).ToString());

                short FD_i = 0;
                double FD_JSHYTD = 0;

                FD_JSHYTD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1HYTD2", 0).Trim());

                for (FD_i = 51; FD_i <= 59; FD_i++)
                {
                    if (FD_JSHYTD <= WG01.TB1AMT[FD_i])
                    {
                        FD_JSHYTD = dataHelpClass.IInt((FD_JSHYTD - WG01.TB1GON[FD_i]) * (WG01.TB1RAT[FD_i] / 100) + WG01.TB1KUM[FD_i], 1);
                        break;
                    }
                }

                //oDS_PH_PY115A.SetValue("U_JS1HYST2", 0, FD_JSHYTD.ToString()); //환산산출세액(2013/01/01 이후)

                //2013/01/01 이후 근속이 있으면
                if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNMON1", 0).Trim()) > 0)
                {
                    oDS_PH_PY115A.SetValue("U_JS1YST01", 0, WK401.YSANTX.ToString()); //연평균산출세액(2013/01/01 이전)
                }
                else
                {
                    oDS_PH_PY115A.SetValue("U_JS1YST01", 0, "0"); //연평균산출세액(2013/01/01 이전)
                }

                oDS_PH_PY115A.SetValue("U_JS1YST02", 0, (FD_JSHYTD / 5).ToString()); //연평균산출세액(2013/01/01 이후)

                //oDS_PH_PY115A.SetValue("U_JS1SAN01", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YST01", 0).Trim()) * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim())).ToString());
                //oDS_PH_PY115A.SetValue("U_JS1SAN02", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YST02", 0).Trim()) * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNYER2", 0).Trim())).ToString());

                oDS_PH_PY115A.SetValue("U_JS1YST", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YST01", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YST02", 0).Trim())).ToString()); //연평균산출세액
                oDS_PH_PY115A.SetValue("U_JS1SAN", 0, (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1SAN01", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1SAN02", 0).Trim())).ToString()); //산출세액
                oDS_PH_PY115A.SetValue("U_JS1STX", 0, WK401.RTXGON.ToString()); //퇴직소득세액공제

                D_GULGAB[0] = dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1SAN", 0).Trim()) - WT_FRNGO1 - WK401.RTXGON, 1);

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //법정외퇴직금
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Compute_TAX_2(WT_RETPA2, WT_RTXYER_2, 0, D_SHWJIG[1], D_SHWYAS[1], D_SHWSUR[1], "BE", 0);

                oDS_PH_PY115A.SetValue("U_JS2RET", 0, WT_RETPA2.ToString()); //퇴직급여총액
                oDS_PH_PY115A.SetValue("U_JS2GON", 0, WK401.RETGON.ToString()); //퇴직소득공제
                oDS_PH_PY115A.SetValue("U_JS2STD", 0, WK401.TAXSTD.ToString()); //퇴직소득과세표준
                oDS_PH_PY115A.SetValue("U_JS2YSD", 0, WK401.YTXSTD.ToString()); //연평균과세표준
                oDS_PH_PY115A.SetValue("U_JS2YST", 0, WK401.YSANTX.ToString()); //연평균산출세액
                oDS_PH_PY115A.SetValue("U_JS2SAN", 0, WK401.SANTAX.ToString()); //산출세액
                oDS_PH_PY115A.SetValue("U_JS2STX", 0, WK401.RTXGON.ToString()); //퇴직소득세액공제

                D_GULGAB[1] = dataHelpClass.IInt(WK401.SANTAX - WT_FRNGO2 - WK401.RTXGON, 1);
                WT_RETPAY = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1RET", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2RET", 0).Trim()); //3.퇴직급여
                WT_RETGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1GON", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2GON", 0).Trim()); //4. 퇴직소득공제(35 = ①특별공제 + ②소득공제)
                WT_TAXSTD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1STD", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2STD", 0).Trim()); //5. 퇴직소득과세표준(36 = 34 - 35)
                WT_YTXSTD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YSD", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2YSD", 0).Trim()); //6. 연평균과세표준(37 = 36 / 근속년수)
                WT_YSANTX = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1YST", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2YST", 0).Trim()); //7. 연평균산출세액(38)
                // WT_SANTAX = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1SAN", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2SAN", 0).Trim()); //8. 산출세액(39 = 38 * 근속년수)   // 20200222 삭제조치(황영수)
                WT_RTXGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1STX", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2STX", 0).Trim()); //9. 퇴직세액공제
                WT_FRNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS1TAX", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2TAX", 0).Trim()); //9. 퇴직세액공제
                WT_GULGAB = D_GULGAB[0] + D_GULGAB[1]; //결정세액
                WT_GULJUM = dataHelpClass.IInt(WT_GULGAB * 0.1, 1); //결정주민세

                // ***************************************************************************
                // 결과값 UPDATE
                // ***************************************************************************
                oDS_PH_PY115A.SetValue("U_RETPAY", 0, WT_RETPAY.ToString()); //퇴직급여총액
                oDS_PH_PY115A.SetValue("U_RETGON", 0, WT_RETGON.ToString()); //퇴직소득공제
                oDS_PH_PY115A.SetValue("U_TAXSTD", 0, WT_TAXSTD.ToString()); //퇴직소득과세표준
                oDS_PH_PY115A.SetValue("U_YTXSTD", 0, WT_YTXSTD.ToString()); //연평균과세표준
                //oDS_PH_PY115A.SetValue("U_YSANTX", 0, WT_YSANTX.ToString()); //연평균산출세액 // 20200222 삭제조치(황영수)
                //oDS_PH_PY115A.SetValue("U_SANTAX", 0, WT_SANTAX.ToString()); //산출세액       // 20200222 삭제조치(황영수)
                oDS_PH_PY115A.SetValue("U_SPCGON", 0, WT_RTXGON.ToString()); //퇴직소득세액공제 
                oDS_PH_PY115A.SetValue("U_TAXGON", 0, WT_FRNGON.ToString()); //외국납부

                if (Convert.ToInt16(codeHelpClass.Left(oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim(), 4)) < 2016)
                {
                    oDS_PH_PY115A.SetValue("U_GULGAB", 0, WT_GULGAB.ToString()); //소득세
                    oDS_PH_PY115A.SetValue("U_GULJUM", 0, WT_GULJUM.ToString()); //주민세

                    if (oDS_PH_PY115A.GetValue("U_MYGACC", 0).Trim() != "" && Convert.ToDouble(oDS_PH_PY115A.GetValue("U_IPGM01", 0).Trim()) != 0)
                    {
                        oDS_PH_PY115A.SetValue("U_GWASEY", 0, (Math.Round(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0).Trim()) * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_IPGM01", 0).Trim()) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()), 0)).ToString());
                        oDS_PH_PY115A.SetValue("U_SEYGAB", 0, oDS_PH_PY115A.GetValue("U_GWASEY", 0));
                        oDS_PH_PY115A.SetValue("U_SEYJUM", 0, (dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SEYGAB", 0).Trim()) * 0.1, 1)).ToString()); //주민세
                    }
                }
                else
                {
                    //2016년 이후 특례적용 산출세액
                    switch (oDS_PH_PY115A.GetValue("U_TAXYEAR", 0).Trim())
                    {
                        case "2016":
                            oDS_PH_PY115A.SetValue("U_SPSANTX", 0, ((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0).Trim()) * 0.8) + (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2SAN03", 0).Trim()) * 0.2)).ToString());
                            break;

                        case "2017":
                            oDS_PH_PY115A.SetValue("U_SPSANTX", 0, ((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0).Trim()) * 0.6) + (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2SAN03", 0).Trim()) * 0.4)).ToString());
                            break;

                        case "2018":
                            oDS_PH_PY115A.SetValue("U_SPSANTX", 0, ((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0).Trim()) * 0.4) + (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2SAN03", 0).Trim()) * 0.6)).ToString());
                            break;

                        case "2019":
                            oDS_PH_PY115A.SetValue("U_SPSANTX", 0, ((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX", 0).Trim()) * 0.2) + (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2SAN03", 0).Trim()) * 0.8)).ToString());
                            break;

                        default:
                            oDS_PH_PY115A.SetValue("U_SPSANTX", 0, oDS_PH_PY115A.GetValue("U_JS2SAN03", 0).Trim());
                            break;
                    }

                    //신고대상 세액
                    oDS_PH_PY115A.SetValue("U_SANTAX1", 0, oDS_PH_PY115A.GetValue("U_SPSANTX", 0).Trim());

                    //과세이연계좌 입력시 이연퇴직 소득세계산
                    if (oDS_PH_PY115A.GetValue("U_MYGACC", 0).Trim() != "" && Convert.ToDouble(oDS_PH_PY115A.GetValue("U_IPGM01", 0)) != 0)
                    {
                        oDS_PH_PY115A.SetValue("U_GWASEY", 0, (Math.Round(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX1", 0).Trim()) * Convert.ToDouble(oDS_PH_PY115A.GetValue("U_IPGM01", 0).Trim()) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()), 0)).ToString());
                        oDS_PH_PY115A.SetValue("U_SEYGAB", 0, oDS_PH_PY115A.GetValue("U_GWASEY", 0).Trim());
                        oDS_PH_PY115A.SetValue("U_SEYJUM", 0, (dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SEYGAB", 0).Trim()) * 0.1, 1)).ToString()); //주민세
                    }

                    oDS_PH_PY115A.SetValue("U_GULGAB", 0, oDS_PH_PY115A.GetValue("U_SANTAX1", 0)); //소득세
                    oDS_PH_PY115A.SetValue("U_GULJUM", 0, (dataHelpClass.IInt(Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SANTAX1", 0)) * 0.1, 1)).ToString()); //주민세

                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("주(현) 퇴직급여가 0 입니다. 퇴직금 계산을 먼저 하시기 바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }

            return returnValue;
        }

        /// <summary>
        /// 근속년수 계산
        /// </summary>
        private void Compute_GNMYER()
        {
            int[] GNMMON = new int[3];
            int[] DUPMON = new int[3];
            int JMMMON;
            int GN1YER;
            int GN2YER;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //법정퇴직
                //1) 주(현)-법정 근무월수
                GNMMON[0] = (oDS_PH_PY115A.GetValue("U_GNSYER", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_GNSYER", 0).Trim())) * 12 + (oDS_PH_PY115A.GetValue("U_GNSMON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_GNSMON", 0).Trim())); //(근속년수 * 12 + 근속월수)

                if ((oDS_PH_PY115A.GetValue("U_GNSDAY", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_GNSDAY", 0).Trim())) > 0)
                {
                    GNMMON[0] = GNMMON[0] + 1;
                }

                DUPMON[0] = (oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim())); //제외일수
                GN1YER = (GNMMON[0] - DUPMON[0]);

                //2) 종(전)-근무월수
                GNMMON[1] = (oDS_PH_PY115A.GetValue("U_JONYER", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_JONYER", 0).Trim())) * 12 + (oDS_PH_PY115A.GetValue("U_JONMON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_JONMON", 0).Trim())); //(종전년수 * 12 + 종전월수)
                DUPMON[1] = (oDS_PH_PY115A.GetValue("U_JEXMON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_JEXMON", 0).Trim())); //종전제외월수

                //3) 중복월수-주(현)과 종(전)일자비교
                JMMMON = (oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim()));

                //4) 세법상 근속년수-법정퇴직
                if (GNMMON[1] > 0)
                {
                    GN1YER = (GNMMON[0] - DUPMON[0] + GNMMON[1] - DUPMON[1] - JMMMON);
                }

                if ((GN1YER % 12) == 0)
                {
                    GN1YER /= 12;
                }
                else
                {
                    GN1YER = Convert.ToInt32(dataHelpClass.IInt(GN1YER / 12, 1)) + 1;
                }

                //법정외 퇴직
                //1) 주(현)-법정외 근무월수
                GNMMON[2] = (oDS_PH_PY115A.GetValue("U_GN2MON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_GN2MON", 0).Trim()));
                DUPMON[2] = (oDS_PH_PY115A.GetValue("U_EX2MON", 0).Trim() == "" ? 0 : Convert.ToInt32(oDS_PH_PY115A.GetValue("U_EX2MON", 0).Trim()));
                //2) 종(전)-근무월수

                //4) 세법상 근속년수-법정외퇴직
                GN2YER = GNMMON[2] - DUPMON[2];
                if ((GN2YER % 12) == 0)
                {
                    GN2YER /= 12;
                }
                else
                {
                    GN2YER = Convert.ToInt32(dataHelpClass.IInt(GN2YER / 12, 1)) + 1;
                }

                //Display
                oDS_PH_PY115A.SetValue("U_GNMMON", 0, GNMMON[0].ToString()); //주(현) 법정 근속월수
                oDS_PH_PY115A.SetValue("U_GNMDAY", 0, GNMMON[1].ToString()); //종(전) 법정 근속월수(필드대체해서사용하는중)
                oDS_PH_PY115A.SetValue("U_GNMYER", 0, GN1YER.ToString()); //세법상의 주현 법정 근속년수
                oDS_PH_PY115A.SetValue("U_GN2YER", 0, GN2YER.ToString()); //세법상의 주현 법정외 근속년수

                oForm.Items.Item("GNMMON").Update();
                oForm.Items.Item("GNMDAY").Update();
                oForm.Items.Item("GNMYER").Update();
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
        /// 퇴직연금 명세 계산
        /// </summary>
        private void Display_ilSiGum()
        {
            double WK_MYNWON;
            double WK_MYNBUL;
            double WK_MYNTOT;
            double WK_MYNGON;
            double WK_YILPAY;
            double WK_JYNWON;
            double WK_JYNBUL;
            double WK_JYNTOT;
            double WK_JYNGON;
            double WK_JYIL01;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //주현 퇴직연금 명세
                WK_MYNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNTOT", 0).Trim()); //총수령액
                WK_MYNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNWON", 0).Trim()); //원리금합계액

                WK_MYNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNBUL", 0).Trim()); //소득자불입액
                WK_MYNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNGON", 0).Trim()); //퇴직연금소득공제액

                //종전 퇴직연금 명세
                WK_JYNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNTOT", 0).Trim()); //총수령액
                WK_JYNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNWON", 0).Trim()); //원리금합계액

                WK_JYNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNBUL", 0).Trim()); //소득자불입액
                WK_JYNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNGON", 0).Trim()); //퇴직연금소득공제액

                //퇴직연금일시금 계산-주현
                if (WK_MYNWON == 0)
                {
                    WK_MYNWON = 1;
                }

                if (WK_JYNWON == 0)
                {
                    WK_JYNWON = 1;
                }

                WK_YILPAY = dataHelpClass.IInt(WK_MYNTOT * (1 - (WK_MYNBUL - WK_MYNGON) / WK_MYNWON), 1);
                WK_JYIL01 = dataHelpClass.IInt(WK_JYNTOT * (1 - (WK_JYNBUL - WK_JYNGON) / WK_JYNWON), 1);

                //총수령액이 없으면 일시금도 0
                if (WK_JYNTOT == 0)
                {
                    WK_JYIL01 = 0;
                }

                if (WK_MYNTOT == 0)
                {
                    WK_YILPAY = 0;
                }

                //Display
                oDS_PH_PY115A.SetValue("U_JYIL01", 0, WK_JYIL01.ToString());
                oForm.Items.Item("JYIL01").Update();

                oDS_PH_PY115A.SetValue("U_YILPAY", 0, WK_YILPAY.ToString());
                oForm.Items.Item("YILPAY").Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 세액환산 명세 계산
        /// </summary>
        /// <param name="HS_TYPE">1-법정퇴직금세액환산명세, 2-법정외퇴직금세액환산명세, 3-주현퇴직금세액환산명세, 4-종전퇴직세액환산명세</param>
        /// <param name="RPTGBN"></param>
        private void Display_TaxWhanSan(short HS_TYPE, string RPTGBN)
        {
            short i;

            double FD_YUNTOT = 0;
            double FD_YUNBUL = 0;
            double FD_YUNGON = 0;
            double FD_YUNWON = 0;
            double FD_RETAMT = 0;
            int FD_GNMYER;
            int FD_GNMMON = 0;

            //Clear
            HS.SHWJIG = 0;
            HS.SHWTIL = 0;
            HS.SHWSUR = 0;
            HS.SHWGON = 0;
            HS.SHWGWA = 0;
            HS.SHWYAG = 0;
            HS.SHWYAS = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //퇴직연금자료
                switch (HS_TYPE)
                {
                    case 1:

                        //지급예상액
                        SHWJIG[0] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1JIG", 0).Trim());
                        SHWJIG[1] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH3JIG", 0).Trim());
                        HS.SHWJIG = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1JIG", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH3JIG", 0).Trim()); //법정지급예상액

                        FD_GNMMON = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMMON", 0).Trim()) - Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim());
                        FD_YUNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNTOT", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNTOT", 0).Trim()); //퇴직연금=총수령액
                        FD_YUNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNBUL", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNBUL", 0).Trim()); //퇴직연금=불입액
                        FD_YUNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNGON", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNGON", 0).Trim()); //퇴직연금=공제액
                        FD_YUNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNWON", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNWON", 0).Trim()); //퇴직연금=원리금합계액

                        //법정퇴직급여=주현+종전
                        if (RPTGBN == "Y")
                        {
                            FD_RETAMT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim());

                            if (oDS_PH_PY115A.GetValue("U_ST2RET", 0).Trim() != "")
                            {
                                FD_RETAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JSUD01", 0).Trim());
                            }
                        }
                        else
                        {
                            FD_RETAMT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JADD01", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1IYO", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2IYO", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH3IYO", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH4IYO", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1GSU", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2GSU", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH3GSU", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH4GSU", 0).Trim()) +
                                        Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2TIL", 0).Trim());
                        }

                        break;

                    case 2:

                        //지급예상액
                        SHWJIG[0] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2JIG", 0).Trim());
                        SHWJIG[1] = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH4JIG", 0).Trim());
                        HS.SHWJIG = SHWJIG[0] + SHWJIG[1]; //지급예상액

                        FD_YUNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNBUL", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNBUL", 0).Trim()); // / 퇴직연금=불입액
                        FD_YUNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNGON", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNGON", 0).Trim()); // / 퇴직연금=공제액
                        FD_YUNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNWON", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNWON", 0).Trim()); // / 퇴직연금=원리금합계액
                        FD_GNMMON = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMMON", 0).Trim()) - Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim());
                        FD_YUNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNTOT", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNTOT", 0).Trim()); // / 퇴직연금=총수령액

                        //법정외퇴직급여=주현+종전
                        if (HS.SHWJIG > 0)
                        {
                            FD_RETAMT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim());

                            if (oDS_PH_PY115A.GetValue("U_ST2RET", 0).Trim() == "")
                            {
                                FD_RETAMT += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JSUD01", 0).Trim());
                            }
                        }
                        else
                        {
                            FD_RETAMT = 0;
                        }

                        break;

                    case 3: //주현

                        //지급예상액
                        HS.SHWJIG = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1JIG", 0));
                        FD_GNMMON = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMMON", 0)) - Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0));
                        FD_YUNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNTOT", 0).Trim()); //퇴직연금=총수령액
                        FD_YUNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNBUL", 0).Trim()); //퇴직연금=불입액
                        FD_YUNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNGON", 0).Trim()); //퇴직연금=공제액
                        FD_YUNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MYNWON", 0).Trim()); //퇴직연금=원리금합계액
                        FD_RETAMT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TJKPAY", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim());

                        break;

                    case 4:

                        //지급예상액
                        HS.SHWJIG = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2JIG", 0).Trim());
                        FD_GNMMON = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMDAY", 0).Trim()) - Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JEXMON", 0).Trim());
                        FD_YUNTOT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNTOT", 0).Trim()); //퇴직연금=총수령액
                        FD_YUNBUL = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNBUL", 0).Trim()); //퇴직연금=불입액
                        FD_YUNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNGON", 0).Trim()); //퇴직연금=공제액
                        FD_YUNWON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYNWON", 0).Trim()); //퇴직연금=원리금합계액
                        FD_RETAMT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JSUD01", 0).Trim());

                        break;
                }

                //세법상의 근속년수
                if ((FD_GNMMON % 12) == 0)
                {
                    FD_GNMYER = FD_GNMMON / 12;
                }
                else
                {
                    FD_GNMYER = Convert.ToInt16(dataHelpClass.IInt(FD_GNMMON / 12, 1)) + 1;
                }

                //******************************************************************************************************************************
                //세액환산계산
                //세액환산명세,퇴직연금일시금지급예상액, 이연금액등이 있을경우 계산됨.
                //******************************************************************************************************************************
                if (HS.SHWJIG > 0)
                {
                    if (RPTGBN == "Y") //명예퇴직
                    {
                        switch (HS_TYPE)
                        {
                            case 1: //법정

                                //22.총일시금(WK_SHWTIL) : 불입액이 있을 경우(지급예상액 * (1-(불입액-소득공제액)/총수령액)),없을 경우(지급예상액과 동일)
                                if (FD_YUNWON != 0)
                                {
                                    HS.SHWTIL = dataHelpClass.IInt(HS.SHWJIG * (1 - (FD_YUNBUL - FD_YUNGON) / FD_YUNWON), 1);
                                }
                                else
                                {
                                    HS.SHWTIL = SHWJIG[0];
                                }

                                break;

                            case 2:
                                //22.총일시금(WK_SHWTIL)
                                HS.SHWTIL = SHWJIG[1];

                                break;
                        }
                    }
                    else
                    {
                        //22.총일시금(WK_SHWTIL) : 불입액이 있을 경우(지급예상액 * (1-(불입액-소득공제액)/총수령액)),없을 경우(지급예상액과 동일)
                        if (FD_YUNTOT != 0)
                        {
                            HS.SHWTIL = dataHelpClass.IInt(HS.SHWJIG * (1 - (FD_YUNBUL - FD_YUNGON) / FD_YUNWON), 1);
                        }
                        else
                        {
                            HS.SHWTIL = HS.SHWJIG;
                        }
                    }

                    //23.수령가능퇴직급여액(WK_SHWSUR)= 총일시금+주현퇴직금+주현명예수당+종전퇴직금+종전명예수당
                    HS.SHWSUR = HS.SHWTIL + FD_RETAMT;

                    //24.환산퇴직소득공제=(45%공제액+근속년수별공제)
                    Compute_TAX_2(HS.SHWSUR, FD_GNMYER, 0, HS.SHWJIG, HS.SHWTIL, HS.SHWSUR, "", 0);
                    HS.SHWGON = WK401.RETGON;

                    //25.환산퇴직소득과세표준(HS.SHWGWA)
                    HS.SHWGWA = WK401.TAXSTD;

                    //26.환산연평균과세표준(HS.SHWYAG)
                    HS.SHWYAG = WK401.YTXSTD;

                    if (HS.SHWYAG < 0)
                    {
                        HS.SHWYAG = 0;
                    }

                    //27.환산연평균산출세액(HS.SHWYAS)
                    HS.SHWYAS = WK401.YSANTX;
                }
                else
                {
                    HS.SHWTIL = 0;
                    HS.SHWSUR = 0;
                    HS.SHWGON = 0;
                    HS.SHWGWA = 0;
                    HS.SHWYAG = 0;
                    HS.SHWYAS = 0;
                }

                //Display
                switch (HS_TYPE)
                {
                    case 1:

                        i = HS_TYPE;

                        oDS_PH_PY115A.SetValue("U_SH" + i + "TIL", 0, HS.SHWTIL.ToString());
                        oDS_PH_PY115A.SetValue("U_SH" + i + "SUR", 0, HS.SHWSUR.ToString());
                        oDS_PH_PY115A.SetValue("U_SH" + i + "GON", 0, HS.SHWGON.ToString());
                        oDS_PH_PY115A.SetValue("U_SH" + i + "GWA", 0, HS.SHWGWA.ToString());
                        oDS_PH_PY115A.SetValue("U_SH" + i + "YAG", 0, HS.SHWYAG.ToString());
                        oDS_PH_PY115A.SetValue("U_SH" + i + "YAS", 0, HS.SHWYAS.ToString());

                        oForm.Items.Item("SH" + i + "TIL").Update();
                        oForm.Items.Item("SH" + i + "SUR").Update();
                        oForm.Items.Item("SH" + i + "GON").Update();
                        oForm.Items.Item("SH" + i + "GWA").Update();
                        oForm.Items.Item("SH" + i + "YAG").Update();
                        oForm.Items.Item("SH" + i + "YAS").Update();

                        break;

                    case 2:

                        i = HS_TYPE;

                        oDS_PH_PY115A.SetValue("U_SH" + i + "TIL", 0, HS.SHWTIL.ToString());

                        oForm.Items.Item("SH" + i + "TIL").Update();

                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 퇴직소득세액  계산
        /// </summary>
        /// <param name="P_RETAMT">대상퇴직금</param>
        /// <param name="P_GNMYER">세법상 근속년수</param>
        /// <param name="P_DUPYER">중복년수</param>
        /// <param name="P_SHWJIG">세액환산명세-퇴직연금지급예상액</param>
        /// <param name="P_SHWYAS">세액환산명세-환산연평균산출세액</param>
        /// <param name="P_SHWSUR">세액환산명세-수령가능퇴직급여액</param>
        /// <param name="HS_TYPE"></param>
        /// <param name="SODAMT"></param>
        private void Compute_TAX_2(double P_RETAMT, int P_GNMYER, short P_DUPYER, double P_SHWJIG, double P_SHWYAS, double P_SHWSUR, string HS_TYPE, double SODAMT)
        {

            short FD_i;
            string FD_RTXGBN_H;
            string FD_RTXGBN_J;
            double FD_YTXSTD;
            double FD_YSANTX = 0;

            //Clear
            WK401.SPCGON = 0;
            WK401.SODGON = 0;
            WK401.RETGON = 0;
            WK401.TAXSTD = 0;
            WK401.YSANTX = 0;
            WK401.YTXSTD = 0;
            WK401.SANTAX = 0;
            WK401.RTXGON = 0;
            WK401.FRNGON = 0;
            WK401.GULGAB = 0;
            WK401.WSGGON = 0; //환산급여별 공제
            
            //double JS2AMT; //환산급여 저장용

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //if (P_RETAMT <= 0)
                //{
                //    return;
                //}

                //세액공제여부(Y/N)
                FD_RTXGBN_H = (oDS_PH_PY115A.GetValue("U_SPCGBN", 0).Trim() == "" ? "N" : oDS_PH_PY115A.GetValue("U_SPCGBN", 0).Trim()); //주현근무지
                FD_RTXGBN_J = (oDS_PH_PY115A.GetValue("U_JS1SPC", 0).Trim() == "" ? "N" : oDS_PH_PY115A.GetValue("U_JS1SPC", 0).Trim()); //종전근무지

                //퇴직소득공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                //① 특별공제: 퇴직급여공제(일반퇴직금+명예퇴직수당(SUDAMT)+단체퇴직수당(BHMAMT))
                //(퇴직급여+단체퇴직보험료+2002년귀속분부터의 명예퇴직수당) * 50%
                //(2001년귀속분까지 명예퇴직수당) * 75%
                WK401.SPCGON = dataHelpClass.IInt(P_RETAMT * (WG01.TB1RAT[91] / 100), 1);

                //② 퇴직소득공제: 근속연수에 따른 공제
                //근속연수에 따라 다음의 초과누진방법으로 공제
                //(근속연수: 5년이하             -           30만원*근속연수
                //           5년초과~10년이하    -  150만원+ 50만원*(근속연수- 5년)
                //           10년초과~20년이상   -  400만원+ 80만원*(근속연수-10년)
                //           20년초과            - 1200만원+120만원*(근속연수-20))
                if (HS_TYPE == "BB")
                {
                    for (FD_i = 93; FD_i <= 96; FD_i++)
                    {
                        if ((P_GNMYER - P_DUPYER) <= WG01.TB1AMT[FD_i])
                        {
                            WK401.SODGON = WG01.TB1KUM[FD_i] + WG01.TB1GON[FD_i] * (P_GNMYER - P_DUPYER - WG01.TB1AMT[FD_i - 1]);
                            break;
                        }
                    }
                }

                WK401.SODGON = dataHelpClass.IInt(WK401.SODGON, 1);

                //환산급여별공제 선행 계산("환산급여" 필드에 값이 바인딩 되지 않은 시점에 변수에 저장하여 "환산급여별공제" 계산)
                //JS2AMT = (P_RETAMT - WK401.SODGON) / Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GNMYER", 0).Trim()) * 12;

                if (Convert.ToInt16(codeHelpClass.Left(oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim(), 4)) >= 2016)
                {
                    //환산급여별 공제 2016년1월 1일 부터 시행
                    if (HS_TYPE == "BB")
                    {
                        for (FD_i = 81; FD_i <= 85; FD_i++)
                        {
                            if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2AMT", 0).Trim()) <= WG01.TB1AMT[FD_i])
                            //if (JS2AMT <= WG01.TB1AMT[FD_i])
                            {
                                WK401.WSGGON = WG01.TB1KUM[FD_i] + ((Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JS2AMT", 0).Trim()) - WG01.TB1GON[FD_i]) * WG01.TB1RAT[FD_i] / 100);
                                //WK401.WSGGON = WG01.TB1KUM[FD_i] + ((JS2AMT - WG01.TB1GON[FD_i]) * WG01.TB1RAT[FD_i] / 100);
                                break;
                            }
                        }
                    }
                }
                // WK401.WSGGON

                //예외처리) 법정외 퇴직

                if (HS_TYPE == "2")
                {
                    //법정퇴직년도공제만큼 차감
                    WK401.SODGON -= SODAMT;
                }
                else if (HS_TYPE == "3")
                {
                    return;
                }

                //(차감)
                if (WK401.SODGON > (P_RETAMT - WK401.SPCGON))
                {
                    WK401.SODGON = P_RETAMT - WK401.SPCGON;
                }
                //퇴직소득공제(22 = ①특별공제 + ②소득공제)
                WK401.RETGON = WK401.SPCGON + WK401.SODGON;
                //퇴직소득과세표준(23 = 21 - 22)
                WK401.TAXSTD = P_RETAMT - WK401.RETGON;
                //연평균과세표준(24 = 23 / 근속년수)
                WK401.YTXSTD = dataHelpClass.IInt(WK401.TAXSTD / P_GNMYER, 1);

                //연평균산출세액(25)
                if (BubJung_YN == "Y" && WT_SHWTOT != 0)
                {
                    //전체 연평균산출세액(법정+법정외)
                    if (HS_TYPE == "2")
                    {
                        FD_YTXSTD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1YAG", 0).Trim()) + WK401.YTXSTD;
                    }
                    else
                    {
                        FD_YTXSTD = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH1YAG", 0).Trim()) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SH2YAG", 0).Trim());
                    }

                    //(소득세 기본세율 적용)
                    for (FD_i = 51; FD_i <= 59; FD_i++)
                    {
                        if (FD_YTXSTD <= WG01.TB1AMT[FD_i])
                        {
                            //2013/12/25 노근용수정
                            FD_YSANTX = dataHelpClass.IInt((FD_YTXSTD - WG01.TB1GON[FD_i]) * (WG01.TB1RAT[FD_i] / 100) + WG01.TB1KUM[FD_i], 1);
                            break;
                        }
                    }

                    if (FD_YSANTX < 0)
                    {
                        FD_YSANTX = 0;
                    }

                    oDS_PH_PY115A.SetValue("U_YSANTX", 0, System.Convert.ToString(FD_YSANTX)); //전체산출세액먼저 셋팅

                    //연평균산출세액=현재연평균과세표준/법정+법정외연평균과세표준* 총산출세액
                    if (FD_YTXSTD != 0)
                    {
                        WK401.YSANTX = Math.Round(WK401.YTXSTD / FD_YTXSTD * FD_YSANTX, 0);
                    }
                }
                else
                    //(소득세 기본세율 적용)
                    for (FD_i = 51; FD_i <= 59; FD_i++)
                    {
                        if (WK401.YTXSTD <= WG01.TB1AMT[FD_i])
                        {
                            //2013/12/25 노근용수정
                            WK401.YSANTX = Math.Round((WK401.YTXSTD - WG01.TB1GON[FD_i]) * (WG01.TB1RAT[FD_i] / 100) + WG01.TB1KUM[FD_i], 0);
                            break;
                        }
                    }

                WK401.YSANTX = dataHelpClass.IInt(WK401.YSANTX, 1);

                if (WK401.YSANTX < 0)
                {
                    WK401.YSANTX = 0;
                }

                //산출세액 (26 = 25 * 근속년수)
                if (P_SHWJIG > 0)
                {
                    if (P_SHWSUR != 0)
                    {
                        WK401.SANTAX = dataHelpClass.IInt(P_SHWYAS * P_GNMYER * P_RETAMT / P_SHWSUR, 1);
                    }
                }
                else
                {
                    WK401.SANTAX = dataHelpClass.IInt(WK401.YSANTX * P_GNMYER, 1);
                }

                //퇴직세액공제(2009년 한시적:근무년수* 24만원한도) =  WK401.TAXGON
                if (FD_RTXGBN_H == "Y")
                {
                    if (P_SHWJIG > 0)
                    {
                        if (WK401.SANTAX > 0)
                        {
                            WK401.RTXGON = dataHelpClass.IInt(WK401.SANTAX * WK401.SANTAX / WK401.SANTAX * 0.3, 1);
                        }
                        else
                        {
                            WK401.RTXGON = 0;
                        }
                    }
                    else
                    {
                        WK401.RTXGON = dataHelpClass.IInt(WK401.SANTAX * 0.3, 1);
                    }

                    if ((P_GNMYER * 240000) < WK401.RTXGON)
                    {
                        WK401.RTXGON = (P_GNMYER * 240000);
                    }
                }
                else
                {
                    WK401.RTXGON = 0;
                }

                //외국납부세액공제
                WK401.FRNGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_TAXGON", 0).Trim()); //외국납부세액

                //결정세액(26)
                WK401.GULGAB = dataHelpClass.IInt(WK401.SANTAX - WK401.FRNGON - WK401.RTXGON, 1);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 퇴직금 생성
        /// </summary>
        /// <returns></returns>
        private bool Create_Data()
        {
            string sQry;
            short errNum = 0;

            int i;
            int j = 0;

            string[] PAYSTD = new string[4];
            string[] PAYEND = new string[4];
            string[] PAYDAY = new string[4];
            double[] PAYAMT = new double[4];
            string[] BNSYMM = new string[12];
            double[] BNSAMT = new double[12];

            double YONCSU;
            string CLTCOD;
            string MSTCOD;
            string RETDAT;
            double KUKJUN;
            double HUGA;
            string JSNGBN = string.Empty;
            string Div; //퇴충계산시: 1, 퇴직금계산시 :2

            bool returnValue = false;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (PSH_Globals.SBO_Application.MessageBox("퇴직금 생성을 하시겠습니까?", 2, "&Yes!", "&No") == 2)
                {
                    returnValue = false;
                }

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                for (i = 0; i < 4; i++)
                {
                    PAYSTD[i] = "";
                    PAYEND[i] = "";
                    PAYDAY[i] = "";
                    PAYAMT[i] = 0;
                }

                for (i = 0; i < 12; i++)
                {
                    BNSYMM[i] = "";
                    BNSAMT[i] = 0;
                }

                YONCSU = 0;

                //자료 조회
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                CLTCOD = dataHelpClass.Get_ReData("U_CLTCOD", "Code", "[@PH_PY001A]", "'" + MSTCOD + "'", "");
                RETDAT = dataHelpClass.ConvertDateType(oDS_PH_PY115A.GetValue("U_ENDRET", 0).ToString().Trim(), "-");

                if (oForm.Items.Item("JSNGBN").Specific.Selected != null)
                {
                    JSNGBN = oForm.Items.Item("JSNGBN").Specific.Selected.Value;
                }

                Div = "2"; //퇴직금계산

                sQry = "Exec PH_PY115_01 '" + MSTCOD + "' , '" + RETDAT + "', '" + FIXTYP + "', '" + ROUNDT + "', '" + RODLEN + "', '" + RETCHK + "','','','" + RETYCH + "', '" + Div + "' ";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount <= 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    if (Convert.ToDouble(oRecordSet.Fields.Item("PAYT04").Value) > 0)
                    {
                        PAYSTD[0] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD1").Value.ToString().Replace("-", ""), 8);
                        PAYEND[0] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED1").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[0] = oRecordSet.Fields.Item("PAYDY1").Value.ToString().Trim();
                        PAYAMT[0] = oRecordSet.Fields.Item("PAYT01").Value;

                        PAYSTD[1] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD2").Value.ToString().Replace("-", ""), 8);
                        PAYEND[1] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED2").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[1] = oRecordSet.Fields.Item("PAYDY2").Value.ToString().Trim();
                        PAYAMT[1] = oRecordSet.Fields.Item("PAYT02").Value;

                        PAYSTD[2] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD3").Value.ToString().Replace("-", ""), 8);
                        PAYEND[2] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED3").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[2] = oRecordSet.Fields.Item("PAYDY3").Value.ToString().Trim(); ;
                        PAYAMT[2] = oRecordSet.Fields.Item("PAYT03").Value;

                        PAYSTD[3] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD4").Value.ToString().Replace("-", ""), 8);
                        PAYEND[3] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED4").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[3] = oRecordSet.Fields.Item("PAYDY4").Value.ToString().Trim();
                        PAYAMT[3] = oRecordSet.Fields.Item("PAYT04").Value;
                    }
                    else
                    {
                        PAYSTD[0] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD1").Value.ToString().Replace("-", ""), 8);
                        PAYEND[0] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED1").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[0] = oRecordSet.Fields.Item("PAYDY1").Value.ToString().Trim();
                        PAYAMT[0] = oRecordSet.Fields.Item("PAYT01").Value;

                        PAYSTD[1] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD2").Value.ToString().Replace("-", ""), 8);
                        PAYEND[1] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED2").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[1] = oRecordSet.Fields.Item("PAYDY2").Value.ToString().Trim();
                        PAYAMT[1] = oRecordSet.Fields.Item("PAYT02").Value;

                        PAYSTD[2] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYSD3").Value.ToString().Replace("-", ""), 8);
                        PAYEND[2] = codeHelpClass.Left(oRecordSet.Fields.Item("PAYED3").Value.ToString().Replace("-", ""), 8);
                        PAYDAY[2] = oRecordSet.Fields.Item("PAYDY3").Value.ToString().Trim();
                        PAYAMT[2] = oRecordSet.Fields.Item("PAYT03").Value;
                    }

                    BNSYMM[0] = oRecordSet.Fields.Item("BNSD01").Value;
                    BNSAMT[0] = oRecordSet.Fields.Item("BNSA01").Value;
                    BNSYMM[1] = oRecordSet.Fields.Item("BNSD02").Value;
                    BNSAMT[1] = oRecordSet.Fields.Item("BNSA02").Value;
                    BNSYMM[2] = oRecordSet.Fields.Item("BNSD03").Value;
                    BNSAMT[2] = oRecordSet.Fields.Item("BNSA03").Value;
                    BNSYMM[3] = oRecordSet.Fields.Item("BNSD04").Value;
                    BNSAMT[3] = oRecordSet.Fields.Item("BNSA04").Value;
                    BNSYMM[4] = oRecordSet.Fields.Item("BNSD05").Value;
                    BNSAMT[4] = oRecordSet.Fields.Item("BNSA05").Value;
                    BNSYMM[5] = oRecordSet.Fields.Item("BNSD06").Value;
                    BNSAMT[5] = oRecordSet.Fields.Item("BNSA06").Value;
                    BNSYMM[6] = oRecordSet.Fields.Item("BNSD07").Value;
                    BNSAMT[6] = oRecordSet.Fields.Item("BNSA07").Value;
                    BNSYMM[7] = oRecordSet.Fields.Item("BNSD08").Value;
                    BNSAMT[7] = oRecordSet.Fields.Item("BNSA08").Value;
                    BNSYMM[8] = oRecordSet.Fields.Item("BNSD09").Value;
                    BNSAMT[8] = oRecordSet.Fields.Item("BNSA09").Value;
                    BNSYMM[9] = oRecordSet.Fields.Item("BNSD10").Value;
                    BNSAMT[9] = oRecordSet.Fields.Item("BNSA10").Value;
                    BNSYMM[10] = oRecordSet.Fields.Item("BNSD11").Value;
                    BNSAMT[10] = oRecordSet.Fields.Item("BNSA11").Value;
                    BNSYMM[11] = oRecordSet.Fields.Item("BNSD12").Value;
                    BNSAMT[11] = oRecordSet.Fields.Item("BNSA12").Value;

                    YONCSU = oRecordSet.Fields.Item("YONCSU").Value; //연차수당
                    KUKJUN = oRecordSet.Fields.Item("KUKJUN").Value; //퇴직전환금
                    HUGA = oRecordSet.Fields.Item("HUGA").Value; //휴가비

                    //퇴직금세액공제(2009년 한시적)
                    if (codeHelpClass.Left(RETDAT, 4) == "2009" && JSNGBN == "1")
                    {
                        oDS_PH_PY115A.SetValue("U_SPCGBN", 0, "Y");
                    }
                    else
                    {
                        oDS_PH_PY115A.SetValue("U_SPCGBN", 0, "N");
                    }

                }

                // ***************************************************************************
                // 결과값 UPDATE
                // ***************************************************************************
                //급여자료
                oDS_PH_PY115A.SetValue("U_PAYSD1", 0, PAYSTD[0]);
                oDS_PH_PY115A.SetValue("U_PAYED1", 0, PAYEND[0]);
                oDS_PH_PY115A.SetValue("U_PAYDY1", 0, PAYDAY[0]);
                oDS_PH_PY115A.SetValue("U_PAYT01", 0, PAYAMT[0].ToString());

                oDS_PH_PY115A.SetValue("U_PAYSD2", 0, PAYSTD[1]);
                oDS_PH_PY115A.SetValue("U_PAYED2", 0, PAYEND[1]);
                oDS_PH_PY115A.SetValue("U_PAYDY2", 0, PAYDAY[1]);
                oDS_PH_PY115A.SetValue("U_PAYT02", 0, PAYAMT[1].ToString());

                oDS_PH_PY115A.SetValue("U_PAYSD3", 0, PAYSTD[2]);
                oDS_PH_PY115A.SetValue("U_PAYED3", 0, PAYEND[2]);
                oDS_PH_PY115A.SetValue("U_PAYDY3", 0, PAYDAY[2]);
                oDS_PH_PY115A.SetValue("U_PAYT03", 0, PAYAMT[2].ToString());

                oDS_PH_PY115A.SetValue("U_PAYSD4", 0, PAYSTD[3]);
                oDS_PH_PY115A.SetValue("U_PAYED4", 0, PAYEND[3]);
                oDS_PH_PY115A.SetValue("U_PAYDY4", 0, PAYDAY[3]);
                oDS_PH_PY115A.SetValue("U_PAYT04", 0, PAYAMT[3].ToString());

                // / 상여자료
                oDS_PH_PY115A.SetValue("U_BNSD01", 0, BNSYMM[0]);
                oDS_PH_PY115A.SetValue("U_BNST01", 0, BNSAMT[0].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD02", 0, BNSYMM[1]);
                oDS_PH_PY115A.SetValue("U_BNST02", 0, BNSAMT[1].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD03", 0, BNSYMM[2]);
                oDS_PH_PY115A.SetValue("U_BNST03", 0, BNSAMT[2].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD04", 0, BNSYMM[3]);
                oDS_PH_PY115A.SetValue("U_BNST04", 0, BNSAMT[3].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD05", 0, BNSYMM[4]);
                oDS_PH_PY115A.SetValue("U_BNST05", 0, BNSAMT[4].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD06", 0, BNSYMM[5]);
                oDS_PH_PY115A.SetValue("U_BNST06", 0, BNSAMT[5].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD07", 0, BNSYMM[6]);
                oDS_PH_PY115A.SetValue("U_BNST07", 0, BNSAMT[6].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD08", 0, BNSYMM[7]);
                oDS_PH_PY115A.SetValue("U_BNST08", 0, BNSAMT[7].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD09", 0, BNSYMM[8]);
                oDS_PH_PY115A.SetValue("U_BNST09", 0, BNSAMT[8].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD10", 0, BNSYMM[9]);
                oDS_PH_PY115A.SetValue("U_BNST10", 0, BNSAMT[9].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD11", 0, BNSYMM[10]);
                oDS_PH_PY115A.SetValue("U_BNST11", 0, BNSAMT[10].ToString());
                oDS_PH_PY115A.SetValue("U_BNSD12", 0, BNSYMM[11]);
                oDS_PH_PY115A.SetValue("U_BNST12", 0, BNSAMT[11].ToString());

                oDS_PH_PY115A.SetValue("U_YONCSU", 0, YONCSU.ToString()); //연차수당

                ////퇴직금 1-퇴직정산일경우 중도정산소득세 가져오기를 선택한 경우(수행되지 않는 로직(주석처리, 2019.12.09 송명규))
                //if (oDS_PH_PY115A.GetValue("U_JSNGBN", 0).Trim() == "1")
                //{
                //    switch (oCLTOPT[0].Trim())
                //    {
                //        case "1":
                //        case "3":
                //            Create_JsnData();
                //            break;
                //    }
                //}

                oDS_PH_PY115A.SetValue("U_GITJA1", 0, KUKJUN.ToString()); //퇴직전환금
                oDS_PH_PY115A.SetValue("U_ExPay2", 0, HUGA.ToString()); //급여내역 조회

                sQry = "Exec PH_PY115_02 '" + MSTCOD + "','" + RETDAT + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oDS_PH_PY115B.InsertRecord(i);
                        oDS_PH_PY115B.Offset = j;
                        oDS_PH_PY115B.SetValue("U_LineNum", i, (i + 1).ToString());
                        oDS_PH_PY115B.SetValue("U_CSUCOD", i, oRecordSet.Fields.Item(0).Value);
                        oDS_PH_PY115B.SetValue("U_CSUNAM", i, oRecordSet.Fields.Item(1).Value);
                        oDS_PH_PY115B.SetValue("U_CSUAMT1", i, oRecordSet.Fields.Item(2).Value);
                        oDS_PH_PY115B.SetValue("U_CSUAMT2", i, oRecordSet.Fields.Item(3).Value);
                        oDS_PH_PY115B.SetValue("U_CSUAMT3", i, oRecordSet.Fields.Item(4).Value);
                        oDS_PH_PY115B.SetValue("U_CSUAMT4", i, oRecordSet.Fields.Item(5).Value);
                        oRecordSet.MoveNext();
                    }
                }

                oMat1.Columns.Item("CSUAMT1").RightJustified = true;
                oMat1.Columns.Item("CSUAMT2").RightJustified = true;
                oMat1.Columns.Item("CSUAMT3").RightJustified = true;
                oMat1.Columns.Item("CSUAMT4").RightJustified = true;
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();

                returnValue = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("퇴직금 생성 자료가 없습니다!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이전 3개월 급여자료가 없습니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// Create_JsnData, Create_Data에서 호출
        /// </summary>
        private void Create_JsnData()
        {
            string sQry;

            string WK_JSNYER;
            double WK_TOTGON;
            double WK_SILJIG;
            double WK_JSNGAB = 0;
            double WK_JSNJUM = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                WK_JSNYER = codeHelpClass.Left(oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim(), 4);

                //자료 조회
                sQry = "  SELECT      ISNULL(U_CHAGAB,0) AS U_CHAGAB,";
                sQry += "             ISNULL(U_CHAJUM,0) AS U_CHAJUM ";
                sQry += " FROM        [@ZPY504H]";
                sQry += " WHERE       U_JSNYER = '" + WK_JSNYER + "'";
                sQry += "             AND U_JSNGBN = '2'";
                sQry += "             AND U_MSTCOD = '" + oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim() + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    WK_JSNGAB = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    WK_JSNJUM = oRecordSet.Fields.Item(1).Value.ToString().Trim();
                }

                //공제총액(갑근세 + 주민세 + 퇴직전환 + 건강보험정산+정산소득세+정산주민세)
                WK_TOTGON = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_CHAGAB", 0)) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_CHAJUM", 0)) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITJA1", 0)) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MEDJSN", 0)) + WK_JSNGAB + WK_JSNJUM;
                //실지급액(퇴직급여액 - 공제총액)
                WK_SILJIG = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_RETPAY", 0)) + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITAA1", 0)) - WK_TOTGON;

                oDS_PH_PY115A.SetValue("U_JSNGAB", 0, WK_JSNGAB.ToString());
                oDS_PH_PY115A.SetValue("U_JSNJUM", 0, WK_JSNJUM.ToString());
                oDS_PH_PY115A.SetValue("U_TOTGON", 0, WK_TOTGON.ToString());
                oDS_PH_PY115A.SetValue("U_SILJIG", 0, WK_SILJIG.ToString());
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 사용안됨(ITEM_PRESSED 에서 MsterChk = true 일 때 호출)
        /// </summary>
        private void MasterUpdate()
        {
            //SAPbobsCOM.Recordset oRecordSet;
            //string sQry;

            try
            {
                PSH_Globals.SBO_Application.MessageBox("계산확정 후 중간정산은 인사마스터의 기산일을, 퇴직은 퇴직일자 와 재직구분을 꼭 수정바랍니다.");
                // / Question
                // //계산후 인사마스터 기산일에 Update 로직
                // // 김영호 과장요청으로 막음, 계산후 직접 인사마스터에 기산일을 수정하기로함
                // // 기산일 체크는 사원입력시 체크함. '2015/05/07 N.G.Y

                // If Sbo_Application.MessageBox("퇴직금 중도정산일자를 사원마스터에 반영하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
                // Exit Sub
                // End If
                // 
                // Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
                // 
                // sQry = "UPDATE [@PH_PY001A] SET U_RETDAT = '" & Val(oDS_PH_PY115A.GetValue("U_ENDRET", 0)) & "'"
                // sQry = sQry & "   WHERE Code = N'" & oForm.Items("MSTCOD").Specific.Value & "'"
                // oRecordSet.DoQuery sQry
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
        /// 퇴직금 계산
        /// </summary>
        /// <returns>성공여부</returns>
        private bool Compute_PAY()
        {
            string sQry;
            string MSTCOD;
            string stdYear;

            string STRDAT;
            string ENDDAT;
            string INPDAT;
            string OUTDAT;
            string TMPDAT;

            //VB6.FixedLengthString STRDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString ENDDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString INPDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString OUTDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString TMPDAT = new VB6.FixedLengthString(8);

            short errNum = 0;
            int iRow;
            double AVRPAY = 0;
            double AVRBNS = 0;
            double AVRYON = 0;
            double AVRDAY = 0;
            double MONPAY = 0;
            double ExPay2;
            double AvExp2 = 0;
            double TOTPAY = 0;
            double TOTBNS = 0;
            double TOTYON;

            short WK_GNSYER;
            short WK_GNSMON;
            short WK_GNSDAY;
            short WK_TGSDAY;
            short WK_JONYER;
            short WK_JONMON;
            short WK_JMMMON;
            string WK_PAYTYP;
            short WK_GNMYER;
            short WK_GNMMON;
            short WK_JM2MON;
            double WK_YERTJK = 0;
            double WK_MONTJK = 0;
            double WK_DAYTJK = 0;
            double WK_TJKPAY = 0;
            double WK_ADDRAT;
            double WK_GITAA2;
            double WK_SUDAMT;
            int X06_YN;
            short WK_GNSDAY_M;
            string SuSilStr;
            double WK_AMOUNT;

            bool returnValue = false;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //세액 및 계산식 다시 불러오기
                BaseSetting();

                //귀속연도
                STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim(); //정산시작일
                ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim(); //정산종료일
                INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0).Trim(); //입사일자
                OUTDAT = oDS_PH_PY115A.GetValue("U_OUTDAT", 0).Trim(); //퇴사일자

                //누진율 적용시
                WK_ADDRAT = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ADDRAT", 0).Trim());
                if (WK_ADDRAT == 0)
                {
                    WK_ADDRAT = 1;
                }

                //급여형태
                WK_PAYTYP = dataHelpClass.Get_ReData("U_PAYTYP", "Code", "[@PH_PY001A]", "'" + oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim() + "'", "");

                //주현 근속년월일
                WK_GNSYER = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNSYER", 0).Trim()); //주(현) 근속년
                WK_GNSMON = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNSMON", 0).Trim()); //주(현) 근속월
                WK_GNSDAY = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNSDAY", 0).Trim()); //주(현) 근속일
                WK_TGSDAY = Convert.ToInt16(oDS_PH_PY115A.GetValue("U_TGSDAY", 0).Trim()); //주(현) 총계산일수

                //일평균 = 년평균임금 * X09 (근속1년미만나머지 일수)/365
                if (WK_GNSMON > 0)
                {
                    //TMPDAT = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Year, WK_GNSYER, (DateTime)VB6.Format(STRDAT.Value, "0000-00-00")), "YYYYMMDD");
                    TMPDAT = Convert.ToDateTime(dataHelpClass.ConvertDateType(STRDAT, "-")).AddYears(WK_GNSYER).ToString("yyyyMMdd"); //일자 변환 확인 및 테스트 필요
                    WK_GNSDAY_M = Convert.ToInt16(dataHelpClass.TermDay(TMPDAT, ENDDAT));
                }
                else
                {
                    WK_GNSDAY_M = WK_GNSDAY;
                }

                //종전 근속연월일
                WK_JONYER = (oDS_PH_PY115A.GetValue("U_JONYER", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JONYER", 0).Trim())); //종(전) 근속년
                WK_JONMON = (oDS_PH_PY115A.GetValue("U_JONMON", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JONMON", 0).Trim())); //종(전) 근속월
                WK_JMMMON = (oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JMMMON", 0).Trim())); //종(전) 중복월수
                WK_JM2MON = (oDS_PH_PY115A.GetValue("U_JM2MON", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_JM2MON", 0).Trim())); //종(전) 중복월수2

                // ***************************************************************************
                // 근무년수가 1년 이상이 되어야 하므로 1미만은 퇴직금계산에서 제외된다
                // ***************************************************************************
                if (WK_GNSYER == 0 && WK_JONYER == 0)
                {
                    //'ErrNum = 1
                    //'GoTo error_Message
                }

                //근무년월일
                if (WK_GNSDAY > 0 || (WK_GNSMON + WK_JONMON - WK_JMMMON) > 0)
                {
                    WK_GNMYER = Convert.ToInt16(WK_GNSYER + WK_JONYER + 1);
                }
                else
                {
                    WK_GNMYER = Convert.ToInt16(WK_GNSYER + WK_JONYER);
                }

                if (WK_GNSDAY > 0)
                {
                    WK_GNMMON = Convert.ToInt16(WK_GNSYER * 12 + WK_GNSMON + 1);
                }
                else
                {
                    WK_GNMMON = Convert.ToInt16(WK_GNSYER * 12 + WK_GNSMON);
                }

                //X06_YN = 0; //해당월의 마지막 일자
                //AVRPAY = 0; //평균임금
                //AVRBNS = 0;
                //AVRYON = 0;
                //AVRDAY = 0;
                //MONPAY = 0;
                //TOTPAY = 0;
                //TOTBNS = 0;
                //TOTYON = 0;

                //최근3개월간 급여계산일수
                AVRDAY += (oDS_PH_PY115A.GetValue("U_PAYDY1", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYDY1", 0).Trim()));
                AVRDAY += (oDS_PH_PY115A.GetValue("U_PAYDY2", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYDY2", 0).Trim()));
                AVRDAY += (oDS_PH_PY115A.GetValue("U_PAYDY3", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYDY3", 0).Trim()));
                AVRDAY += (oDS_PH_PY115A.GetValue("U_PAYDY4", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYDY4", 0).Trim()));
                //최근3개월간 급여
                TOTPAY += (oDS_PH_PY115A.GetValue("U_PAYT01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYT01", 0).Trim()));
                TOTPAY += (oDS_PH_PY115A.GetValue("U_PAYT02", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYT02", 0).Trim()));
                TOTPAY += (oDS_PH_PY115A.GetValue("U_PAYT03", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYT03", 0).Trim()));
                TOTPAY += (oDS_PH_PY115A.GetValue("U_PAYT04", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_PAYT04", 0).Trim()));
                //최근1년간 상여
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST01", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST01", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST02", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST02", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST03", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST03", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST04", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST04", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST05", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST05", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST06", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST06", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST07", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST07", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST08", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST08", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST09", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST09", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST10", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST10", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST11", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST11", 0).Trim()));
                TOTBNS += (oDS_PH_PY115A.GetValue("U_BNST12", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_BNST12", 0).Trim()));

                TOTYON = (oDS_PH_PY115A.GetValue("U_YONCSU", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_YONCSU", 0).Trim()));

                //평균체력단련비
                ExPay2 = (oDS_PH_PY115A.GetValue("U_ExPay2", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ExPay2", 0).Trim()));

                //임원일경우 누진 근속년 적용
                short GNSYER = 0;
                short GNSMON = 0;
                short GNSDAY = 0;
                short ADDRAT = 0;
                string NUJINYN = string.Empty;

                sQry = "EXEC PH_PY115_03 '" + oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim() + "','" + oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim() + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    GNSYER = oRecordSet.Fields.Item("YY").Value;
                    GNSMON = oRecordSet.Fields.Item("MM").Value;
                    GNSDAY = 0;
                    NUJINYN = oRecordSet.Fields.Item("NUJINYN").Value;
                }
                else
                {
                    NUJINYN = "N";
                }

                if (NUJINYN == "Y")
                {
                    ADDRAT = 1;
                }

                //퇴직급여 계산하기
                for (iRow = 0; iRow < 9; iRow++)
                {
                    if (WK_SILSIL[iRow].ToString().Trim() != "")
                    {
                        SuSilStr = WK_SILSIL[iRow];
                        SuSilStr = SuSilStr.Replace("P01", TOTPAY.ToString()); //총급여 //Strings.Replace(SuSilStr, "P01", System.Convert.ToString(TOTPAY));
                        SuSilStr = SuSilStr.Replace("P02", TOTBNS.ToString()); //총상여 //Strings.Replace(SuSilStr, "P02", System.Convert.ToString(TOTBNS)); 
                        SuSilStr = SuSilStr.Replace("P03", TOTYON.ToString()); //총연차 //Strings.Replace(SuSilStr, "P03", System.Convert.ToString(TOTYON)); 
                        SuSilStr = SuSilStr.Replace("A01", AVRPAY.ToString()); //평균급여 //Strings.Replace(SuSilStr, "A01", System.Convert.ToString(AVRPAY)); 
                        SuSilStr = SuSilStr.Replace("A02", AVRBNS.ToString()); //평균상여 //Strings.Replace(SuSilStr, "A02", System.Convert.ToString(AVRBNS)); 
                        SuSilStr = SuSilStr.Replace("A03", AVRYON.ToString()); //평균연차 //Strings.Replace(SuSilStr, "A03", System.Convert.ToString(AVRYON)); 
                        SuSilStr = SuSilStr.Replace("P04", ExPay2.ToString()); //총체력단련비 //Strings.Replace(SuSilStr, "P04", System.Convert.ToString(ExPay2)); 
                        SuSilStr = SuSilStr.Replace("A04", AvExp2.ToString()); //평균체력단련비 //Strings.Replace(SuSilStr, "A04", AvExp2); 
                        SuSilStr = SuSilStr.Replace("A05", MONPAY.ToString("###.0")); //평균임금 //Replace(SuSilStr, "A05", VB6.Format(MONPAY, "###.0")); 
                        SuSilStr = SuSilStr.Replace("A06", WK_YERTJK.ToString()); //연퇴직금 //Strings.Replace(SuSilStr, "A06", System.Convert.ToString(WK_YERTJK)); 
                        SuSilStr = SuSilStr.Replace("A07", WK_MONTJK.ToString()); //월퇴직금 //Strings.Replace(SuSilStr, "A07", System.Convert.ToString(WK_MONTJK)); 
                        SuSilStr = SuSilStr.Replace("A08", WK_DAYTJK.ToString()); //일퇴직금 //Strings.Replace(SuSilStr, "A08", System.Convert.ToString(WK_DAYTJK)); 
                        SuSilStr = SuSilStr.Replace("A09", WK_TJKPAY.ToString()); //퇴직금계 //Strings.Replace(SuSilStr, "A09", System.Convert.ToString(WK_TJKPAY)); 
                        SuSilStr = SuSilStr.Replace("X01", AVRDAY.ToString()); //급여기간일수 //Strings.Replace(SuSilStr, "X01", System.Convert.ToString(AVRDAY)); 
                        SuSilStr = SuSilStr.Replace("X02", WK_TGSDAY.ToString()); //총근속일수 //Strings.Replace(SuSilStr, "X02", System.Convert.ToString(WK_TGSDAY)); 

                        if (NUJINYN == "N")
                        {
                            SuSilStr = SuSilStr.Replace("X03", WK_GNSYER.ToString()); //근속년수 //Strings.Replace(SuSilStr, "X03", System.Convert.ToString(WK_GNSYER)); 
                            SuSilStr = SuSilStr.Replace("X04", WK_GNSMON.ToString()); //근속월수 //Strings.Replace(SuSilStr, "X04", System.Convert.ToString(WK_GNSMON)); 
                            SuSilStr = SuSilStr.Replace("X05", WK_GNSDAY.ToString()); //근속일수 //Strings.Replace(SuSilStr, "X05", System.Convert.ToString(WK_GNSDAY)); 
                        }
                        else
                        {
                            SuSilStr = SuSilStr.Replace("X03", GNSYER.ToString()); //근속년수 //Strings.Replace(SuSilStr, "X03", System.Convert.ToString(GNSYER)); 
                            SuSilStr = SuSilStr.Replace("X04", GNSMON.ToString()); //근속월수 //Strings.Replace(SuSilStr, "X04", System.Convert.ToString(GNSMON)); 
                            SuSilStr = SuSilStr.Replace("X05", GNSDAY.ToString()); //근속일수 //Strings.Replace(SuSilStr, "X05", System.Convert.ToString(GNSDAY)); 
                        }

                        //X06_YN = Strings.InStr(SuSilStr, "X06");
                        X06_YN = SuSilStr.IndexOf("X06") + 1; //VB의 InStr은 0부터 리턴하지만, C#의 IndexOf는 -1부터 리턴하므로 1을 더해줌

                        if (NUJINYN == "N")
                        {
                            SuSilStr = SuSilStr.Replace("X06", WK_ADDRAT.ToString()); //누진개월(율) //Strings.Replace(SuSilStr, "X06", System.Convert.ToString(WK_ADDRAT)); 
                        }
                        else
                        {
                            SuSilStr = SuSilStr.Replace("X06", ADDRAT.ToString()); //누진개월(율) //Strings.Replace(SuSilStr, "X06", System.Convert.ToString(ADDRAT));
                        }

                        SuSilStr = SuSilStr.Replace("X07", WK_TGSDAY.ToString()); //근속년수계산(2) //Strings.Replace(SuSilStr, "X07", System.Convert.ToString(WK_TGSDAY)); 
                        SuSilStr = SuSilStr.Replace("X08", WK_PAYTYP.ToString()); //급여형태 //Strings.Replace(SuSilStr, "X08", WK_PAYTYP); 
                        SuSilStr = SuSilStr.Replace("X90", WK_GNSDAY_M.ToString()); //근속일수(1년미만 일수) //Strings.Replace(SuSilStr, "X09", System.Convert.ToString(WK_GNSDAY_M));

                        switch (iRow) //iRow가 0부터 시작하므로 VB.6.0 Code 보다 1씩 감소
                        {
                            case 4:
                                oDS_PH_PY115A.SetValue("U_AVRSIL", 0, SuSilStr);
                                break;

                            case 5:
                                oDS_PH_PY115A.SetValue("U_YERSIL", 0, SuSilStr);
                                break;

                            case 6:
                                oDS_PH_PY115A.SetValue("U_MONSIL", 0, SuSilStr);
                                break;

                            case 7:
                                oDS_PH_PY115A.SetValue("U_DAYSIL", 0, SuSilStr);
                                break;
                        }

                        //금액산출
                        sQry = "SELECT 1.0 * " + SuSilStr;
                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.RecordCount != 0)
                        {
                            WK_AMOUNT = dataHelpClass.RInt(Convert.ToDouble(oRecordSet.Fields.Item(0).Value), WK_LENGTH[iRow], WK_ROUNDT[iRow]);
                        }
                        else
                        {
                            WK_AMOUNT = 0;
                        }

                        switch (iRow) //iRow가 0부터 시작하므로 VB.6.0 Code 보다 1씩 감소
                        {
                            case 0:
                                AVRPAY = WK_AMOUNT;
                                break;

                            case 1:
                                AVRBNS = WK_AMOUNT;
                                break;

                            case 2:
                                AVRYON = WK_AMOUNT;
                                break;

                            case 3:
                                AvExp2 = WK_AMOUNT;
                                break;

                            case 4:
                                MONPAY = WK_AMOUNT;
                                break;

                            case 5:
                                WK_YERTJK = WK_AMOUNT;
                                break;

                            case 6:
                                WK_MONTJK = WK_AMOUNT;
                                break;

                            case 7:
                                WK_DAYTJK = WK_AMOUNT;
                                break;

                            case 8: //퇴직금 = 임원누진율사용과 누진적용계산식이 없을경우 누진율자동으로 아니면 수식그대로 사용.

                                if (RATUSE.Trim() == "Y" && X06_YN == 0)
                                {
                                    if (NUJINYN == "N")
                                    {
                                        WK_TJKPAY = WK_AMOUNT * WK_ADDRAT; //중복누진율적용
                                    }
                                    else
                                    {
                                        WK_TJKPAY = WK_AMOUNT * ADDRAT;
                                    }
                                }
                                else
                                {
                                    WK_TJKPAY = WK_AMOUNT;
                                }

                                break;
                        }
                    }
                }

                //기타수당2=미도래연차수당등..
                WK_GITAA2 = (oDS_PH_PY115A.GetValue("U_GITAA2", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITAA2", 0).Trim()));
                WK_SUDAMT = (oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim() == "" ? 0 : Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SUDAMT", 0).Trim()));
                WK_TJKPAY += WK_GITAA2;
                //기타수당1=퇴직금계산에는 제외하고 퇴직금산정서에 포함만 하는 금액임.

                //***************************************************************************
                //결과값 UPDATE
                //***************************************************************************
                oForm.Freeze(true);
                oDS_PH_PY115A.SetValue("U_GNMYER", 0, WK_GNMYER.ToString());
                oDS_PH_PY115A.SetValue("U_GNMMON", 0, WK_GNMMON.ToString());
                oDS_PH_PY115A.SetValue("U_ADDRAT", 0, WK_ADDRAT.ToString());
                oDS_PH_PY115A.SetValue("U_AVRPAY", 0, AVRPAY.ToString()); //평균급여
                oDS_PH_PY115A.SetValue("U_AVRBNS", 0, AVRBNS.ToString()); //평균상여
                oDS_PH_PY115A.SetValue("U_AVRYON", 0, AVRYON.ToString()); //평균연차
                oDS_PH_PY115A.SetValue("U_AvExp2", 0, AvExp2.ToString()); //평균체력단련비
                oDS_PH_PY115A.SetValue("U_MONPAY", 0, MONPAY.ToString()); //평균임금
                oDS_PH_PY115A.SetValue("U_YERTJK", 0, WK_YERTJK.ToString()); //년퇴직금
                oDS_PH_PY115A.SetValue("U_MONTJK", 0, WK_MONTJK.ToString()); //월퇴직금
                oDS_PH_PY115A.SetValue("U_DAYTJK", 0, WK_DAYTJK.ToString()); //일퇴직금
                oDS_PH_PY115A.SetValue("U_TJKPAY", 0, WK_TJKPAY.ToString()); //퇴직금

                stdYear = codeHelpClass.Left(ENDDAT, 4);
                MSTCOD = oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim();
                sQry = "Select b.U_Amt From [@PH_PY125A] a inner Join [@PH_PY125B] b On a.Code = b.Code And a.U_YEAR = '" + stdYear + "' And b.U_MSTCOD = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    oDS_PH_PY115A.SetValue("U_TJKPAY01", 0, oRecordSet.Fields.Item(0).Value); //퇴직연금액
                    oDS_PH_PY115A.SetValue("U_TJKPAY02", 0, WK_TJKPAY - Convert.ToDouble(oRecordSet.Fields.Item(0).Value)); //퇴직연금액
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근속년수가 1년미만이므로, 퇴직금계산이 불가합니다!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("퇴직급여총액이 0입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                
                oForm.Freeze(false);
                oForm.Update();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 근속연월일, 총근무일수구함
        /// </summary>
        private void Compute_GNSDAY()
        {
            //VB6.FixedLengthString STRDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString ENDDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString INPDAT = new VB6.FixedLengthString(8);
            //VB6.FixedLengthString OUTDAT = new VB6.FixedLengthString(8);

            string STRDAT;
            string ENDDAT;
            string INPDAT;
            string OUTDAT;
            short WK_GNSYER;
            short WK_GNSMON;
            short WK_GNSDAY;
            short WK_TGSDAY;
            short WK_GN1YER;
            short WK_GN1MON;
            short WK_EXPMON;
            short WK_GN2YER = 0;
            short WK_GN2MON = 0;
            short WK_EX2MON;
            short errNum = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //귀속연도
                STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim(); //정산시작일
                ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim(); //정산종료일
                INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0).Trim(); //입사일
                OUTDAT = oDS_PH_PY115A.GetValue("U_OUTDAT", 0).Trim(); //퇴사일

                if (STRDAT == "" && ENDDAT == "")
                {
                    errNum = 1;
                    throw new Exception();
                }

                //지급조서 귀속기간 재설정
                if (oDS_PH_PY115A.GetValue("U_STRINT", 0).Trim() == "")
                {
                    if (codeHelpClass.Mid(STRDAT, 0, 4) == codeHelpClass.Mid(ENDDAT, 0, 4))
                    {
                        oDS_PH_PY115A.SetValue("U_STRINT", 0, STRDAT);

                        if (codeHelpClass.Mid(oDS_PH_PY115A.GetValue("U_JOTDAT", 0).Trim(), 0, 4) == codeHelpClass.Mid(ENDDAT, 0, 4))
                        {
                            oDS_PH_PY115A.SetValue("U_STRINT", 0, codeHelpClass.Mid(ENDDAT, 0, 4) + "0101");
                        }
                    }
                    else
                    {
                        oDS_PH_PY115A.SetValue("U_STRINT", 0, codeHelpClass.Mid(ENDDAT, 0, 4) + "0101");
                    }
                }

                if (oDS_PH_PY115A.GetValue("U_ENDINT", 0).Trim() == "")
                {
                    oDS_PH_PY115A.SetValue("U_ENDINT", 0, ENDDAT);
                }

                //1. 법정퇴직금
                dataHelpClass.Term2(STRDAT, ENDDAT);

                //1.1) 근속연월일
                WK_GNSYER = PSH_Globals.ZPAY_GBL_GNSYER;
                WK_GNSMON = PSH_Globals.ZPAY_GBL_GNSMON;
                WK_GNSDAY = PSH_Globals.ZPAY_GBL_GNSDAY;
                WK_TGSDAY = Convert.ToInt16(dataHelpClass.TermDay(STRDAT, ENDDAT));

                //if (Strings.InStr(1, WK_SILSIL[7], "X07") > 0)
                if (WK_SILSIL[6].IndexOf("X07") > 0)
                {
                    WK_TGSDAY = Convert.ToInt16((WK_GNSYER * 360) + (WK_GNSMON * 30) + WK_GNSDAY + 1);
                }

                //1.2) 법정퇴직금 근속월
                WK_GN1MON = Convert.ToInt16(WK_GNSYER * 12 + WK_GNSMON);
                if (WK_GNSDAY > 0)
                {
                    WK_GN1MON = Convert.ToInt16(WK_GN1MON + 1);
                }

                //1.2) 법정퇴직금 제외월
                WK_EXPMON = (oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim()));

                //1.3) 법정퇴직금 근속년
                WK_GN1YER = Convert.ToInt16(WK_GN1MON - WK_EXPMON);
                if ((WK_GN1YER % 12) == 0)
                {
                    WK_GN1YER = Convert.ToInt16(dataHelpClass.IInt(Convert.ToDouble(WK_GN1MON - WK_EXPMON) / 12, 1));
                }
                else
                {
                    WK_GN1YER = Convert.ToInt16(dataHelpClass.IInt(Convert.ToDouble(WK_GN1MON - WK_EXPMON) / 12 + 1, 1));
                }

                //2. 법정이외퇴직금
                if (ENDDAT != "")
                {
                    if (oDS_PH_PY115A.GetValue("U_JINDAT", 0).Trim() == INPDAT)
                    {
                        dataHelpClass.Term2(STRDAT, ENDDAT);
                    }
                    else
                    {
                        dataHelpClass.Term2(INPDAT, ENDDAT);
                    }

                    WK_GN2MON = Convert.ToInt16(PSH_Globals.ZPAY_GBL_GNSYER * 12 + PSH_Globals.ZPAY_GBL_GNSMON);

                    if (PSH_Globals.ZPAY_GBL_GNSDAY > 0)
                    {
                        WK_GN2MON = Convert.ToInt16(WK_GN2MON + 1);
                    }

                    //1.2) 법정퇴직금 제외월
                    WK_EX2MON = (oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim() == "" ? Convert.ToInt16(0) : Convert.ToInt16(oDS_PH_PY115A.GetValue("U_EXPMON", 0).Trim()));
                    WK_GN2YER = Convert.ToInt16(dataHelpClass.IInt(Convert.ToDouble(WK_GN2MON - WK_EX2MON) / 12, 1));

                    if ((WK_GN2YER % 12) == 0)
                    {
                        WK_GN2YER = Convert.ToInt16(dataHelpClass.IInt(Convert.ToDouble(WK_GN2MON - WK_EX2MON) / 12, 1));
                    }
                    else
                    {
                        WK_GN2YER = Convert.ToInt16(dataHelpClass.IInt(Convert.ToDouble(WK_GN2MON - WK_EX2MON) / 12 + 1, 1));
                    }
                }

                //Display
                oDS_PH_PY115A.SetValue("U_GNSYER", 0, WK_GNSYER.ToString()); //근속년
                oDS_PH_PY115A.SetValue("U_GNSMON", 0, WK_GNSMON.ToString()); //근속월
                oDS_PH_PY115A.SetValue("U_GNSDAY", 0, WK_GNSDAY.ToString()); //근속일
                oDS_PH_PY115A.SetValue("U_TGSDAY", 0, WK_TGSDAY.ToString()); //총근속일
                oDS_PH_PY115A.SetValue("U_GNMMON", 0, WK_GN1MON.ToString()); //법정월수
                oDS_PH_PY115A.SetValue("U_GNMYER", 0, WK_GN1YER.ToString()); //세법상의근속년수
                oDS_PH_PY115A.SetValue("U_GN2MON", 0, WK_GN2MON.ToString()); //법정외월수
                oDS_PH_PY115A.SetValue("U_GN2YER", 0, WK_GN2YER.ToString()); //법정외년수

                if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()) <= 20121231)
                {
                    oDS_PH_PY115A.SetValue("U_GNMON1", 0, "0"); //근속월
                    oDS_PH_PY115A.SetValue("U_GNYER1", 0, "0"); //근속년
                    ////2012년12월31 이전 근속월수
                    //STRDAT = oDS_PH_PY115A.GetValue("U_DSDAT1", 0).Trim(); //정산시작일
                    //ENDDAT = oDS_PH_PY115A.GetValue("U_DEDAT1", 0).Trim(); //정산종료일

                    //dataHelpClass.Term2(STRDAT, ENDDAT);

                    //WK_GNSYER = PSH_Globals.ZPAY_GBL_GNSYER;
                    //WK_GNSMON = PSH_Globals.ZPAY_GBL_GNSMON;
                    //WK_GNSDAY = PSH_Globals.ZPAY_GBL_GNSDAY;

                    //WK_GNSMON = Convert.ToInt16(WK_GNSYER * 12 + WK_GNSMON);

                    //if (WK_GNSDAY > 0)
                    //{
                    //    WK_GNSMON = 0;// Convert.ToInt16(WK_GNSMON + 1);
                    //}

                    //oDS_PH_PY115A.SetValue("U_GNMON1", 0, WK_GNSMON.ToString()); //근속월

                    //if ((WK_GNSMON % 12) == 0)
                    //{
                    //    WK_GN2YER = 0;//Convert.ToInt16(dataHelpClass.IInt(WK_GNSMON / 12, 1));
                    //}
                    //else
                    //{
                    //    WK_GN2YER = 0;//Convert.ToInt16(dataHelpClass.IInt(WK_GNSMON / 12 + 1, 1));
                    //}   

                    //oDS_PH_PY115A.SetValue("U_GNYER1", 0, WK_GN2YER.ToString()); //근속년
                }
                else
                {
                    oDS_PH_PY115A.SetValue("U_GNMON1", 0, "0"); //근속월
                    oDS_PH_PY115A.SetValue("U_GNYER1", 0, "0"); //근속년
                }

                //2012년12월31 이후 근속월수
                STRDAT = oDS_PH_PY115A.GetValue("U_DSDAT2", 0).Trim(); //정산시작일
                ENDDAT = oDS_PH_PY115A.GetValue("U_DEDAT2", 0).Trim(); //정산종료일

                dataHelpClass.Term2(STRDAT, ENDDAT);

                WK_GNSYER = PSH_Globals.ZPAY_GBL_GNSYER;
                WK_GNSMON = PSH_Globals.ZPAY_GBL_GNSMON;
                WK_GNSDAY = PSH_Globals.ZPAY_GBL_GNSDAY;

                WK_GNSMON = Convert.ToInt16(WK_GNSYER * 12 + WK_GNSMON);

                if (WK_GNSDAY > 0)
                {
                    WK_GNSMON = Convert.ToInt16(WK_GNSMON + 1);
                }

                oDS_PH_PY115A.SetValue("U_GNMON2", 0, WK_GNSMON.ToString()); //근속월

                WK_GN2YER = Convert.ToInt16(Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNMYER", 0).Trim()) - Convert.ToInt16(oDS_PH_PY115A.GetValue("U_GNYER1", 0).Trim()));

                oDS_PH_PY115A.SetValue("U_GNYER2", 0, WK_GN2YER.ToString()); //근속년
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    //아무실행 없이 메소드 종료
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 퇴직소득세액 계산
        /// 퇴직금 + 세액 계산 -> 세액만 계산
        /// 중도정산년도와 정산년도종년년도가 다르며, 퇴직명예수당등 법정외금액이 있을경우
        /// </summary>
        /// <returns></returns>
        private bool Compute_TAX_3()
        {
            double WK_TOTGON;
            double WK_SILJIG;
            double WK_CHAGAB;
            double WK_CHAJUM;
            double WK_RETPAY;

            bool returnValue = false;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //퇴직급여계
                WK_RETPAY = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_RETPAY", 0).Trim());
                //차감세액 : 결정세액 - 종전납부세액
                WK_CHAGAB = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GULGAB", 0).Trim()) - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JONGAB", 0).Trim()) - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SEYGAB", 0).Trim());
                WK_CHAJUM = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GULJUM", 0).Trim()) - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JONJUM", 0).Trim()) - Convert.ToDouble(oDS_PH_PY115A.GetValue("U_SEYJUM", 0).Trim());
                WK_CHAGAB = dataHelpClass.IInt(WK_CHAGAB, 10);
                WK_CHAJUM = dataHelpClass.IInt(WK_CHAJUM, 10);

                //공제총액(갑근세 + 특별세 + 주민세 + 퇴직전환 + 기타공제+ 건강보험정산 + 외국납부세액)
                WK_TOTGON = WK_CHAGAB + WK_CHAJUM;
                WK_TOTGON += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITJA1", 0).Trim()); //국민연금전환금
                WK_TOTGON += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_MEDJSN", 0).Trim()); //건강보험정산
                WK_TOTGON += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JSNGAB", 0).Trim()); //중도정산소득세
                WK_TOTGON += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JSNJUM", 0).Trim()); //중도정산주민세
                WK_TOTGON += Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITJA2", 0).Trim()); //기타공제

                //실지급액(퇴직급여액 - 공제총액)
                WK_SILJIG = dataHelpClass.IInt(WK_RETPAY + Convert.ToDouble(oDS_PH_PY115A.GetValue("U_GITAA1", 0)) - WK_TOTGON, 1);

                //***************************************************************************
                //결과값 UPDATE
                //***************************************************************************
                oDS_PH_PY115A.SetValue("U_CHAGAB", 0, WK_CHAGAB.ToString()); //차감소득세
                oDS_PH_PY115A.SetValue("U_CHAJUM", 0, WK_CHAJUM.ToString()); //차감주민세
                oDS_PH_PY115A.SetValue("U_TOTGON", 0, WK_TOTGON.ToString()); //공제총계
                oDS_PH_PY115A.SetValue("U_SILJIG", 0, WK_SILJIG.ToString()); //실지급액

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 임원누진율
        /// </summary>
        private void Compute_APPRAT()
        {
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT      ISNULL(T0.U_ADDRAT,1)";
                sQry += " FROM        [@PH_PY114C] T0";
                sQry += "             INNER JOIN";
                sQry += "             [OHPS] T1";
                sQry += "                 ON T0.U_MSTSTP = T1.posID";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY001A] T2";
                sQry += "                 ON T2.U_Position = T1.POSID";
                sQry += " WHERE       T0.Code = '" + oCode + "'";
                sQry += "             AND T2.Code = '" + oDS_PH_PY115A.GetValue("U_MSTCOD", 0).Trim() + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    oDS_PH_PY115A.SetValue("U_ADDRAT", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
                else
                {
                    oDS_PH_PY115A.SetValue("U_ADDRAT", 0, "1");
                }
                oForm.Items.Item("ADDRAT").Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// FlushToItemValue
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        private void FlushToItemValue(string oUID, int oRow)
        {
            
            string STRDAT;
            string ENDDAT;
            short WK_JONYER;
            short WK_JONMON;
            short WK_GNMDAY;
            short WK_DUPMON;
            double JRET01;
            double JYIL03;
            double JTOT01;
            double JADD01;
            double JYIL04;
            double JTOT02;

            ZPAY_g_EmpID oMast; //= new ZPAY_g_EmpID();

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "MSTCOD": //사원번호

                        if (oForm.Items.Item(oUID).Specific.Value == "")
                        {
                            oDS_PH_PY115A.SetValue("U_MSTCOD", 0, "");
                            oDS_PH_PY115A.SetValue("U_MSTNAM", 0, "");
                            oDS_PH_PY115A.SetValue("U_EmpID", 0, "");
                            oDS_PH_PY115A.SetValue("U_INPDAT", 0, "");
                            oDS_PH_PY115A.SetValue("U_OUTDAT", 0, "");
                            oDS_PH_PY115A.SetValue("U_STRRET", 0, "");
                            oDS_PH_PY115A.SetValue("U_ENDRET", 0, "");

                            oDS_PH_PY115A.SetValue("U_DSDAT1", 0, "");
                            oDS_PH_PY115A.SetValue("U_DEDAT1", 0, "");
                            oDS_PH_PY115A.SetValue("U_DSDAT2", 0, "");
                            oDS_PH_PY115A.SetValue("U_DEDAT2", 0, "");
                        }
                        else
                        {
                            oDS_PH_PY115A.SetValue("U_MSTCOD", 0, oForm.Items.Item(oUID).Specific.Value.ToString().ToUpper());
                            oMast = dataHelpClass.Get_EmpID_InFo(oForm.Items.Item(oUID).Specific.Value);
                            oDS_PH_PY115A.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
                            oDS_PH_PY115A.SetValue("U_EmpID", 0, oMast.EmpID);

                            oDS_PH_PY115A.SetValue("U_INPDAT", 0, (GNSGBN == "1" ? oMast.GRPDAT : oMast.StartDate));
                            oDS_PH_PY115A.SetValue("U_OUTDAT", 0, oMast.TermDate);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                //정산시작일자
                                if (oMast.RETDAT.Trim() != oMast.ENDRET.Trim())
                                {
                                    PSH_Globals.SBO_Application.MessageBox("중간정산일과 인사마스터의 기산일을 확인바랍니다.");
                                }

                                if (oMast.RETDAT.Trim() == "" || codeHelpClass.Left(oMast.RETDAT.Trim(), 4) == "1900")
                                {
                                    oDS_PH_PY115A.SetValue("U_STRRET", 0, (GNSGBN == "1" ? oMast.GRPDAT : oMast.StartDate));

                                    //2012년12월31일 이전 기산일, 퇴사일
                                    if (Convert.ToDouble((GNSGBN == "1" ? oMast.GRPDAT : oMast.StartDate)) <= 20121231)
                                    {
                                        oDS_PH_PY115A.SetValue("U_DSDAT2", 0, (GNSGBN == "1" ? oMast.GRPDAT : oMast.StartDate));
                                        //      oDS_PH_PY115A.SetValue("U_DEDAT1", 0, "20121231");
                                    }
                                }
                                else
                                {
                                    //STRDAT = System.Convert.ToString(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 0, (DateTime)VB6.Format(oMast.RETDAT, "0000-00-00")));
                                    //oDS_PH_PY115A.setValue("U_STRRET", 0, Strings.Replace(STRDAT, "-", "")); //왜 변수에 저장했다가 다시 변환("-" > "")하는지 모르겠음, 변환단계 없앰(2019.12.02 송명규)

                                    oDS_PH_PY115A.SetValue("U_STRRET", 0, oMast.RETDAT);

                                    if (Convert.ToDouble(oMast.RETDAT) <= 20121231)
                                    {
                                        //2012년12월31일 이전 기산일, 퇴사일

                                        oDS_PH_PY115A.SetValue("U_DSDAT1", 0, "");
                                        oDS_PH_PY115A.SetValue("U_DEDAT1", 0, "");

                                        //oDS_PH_PY115A.SetValue("U_DSDAT1", 0, oMast.RETDAT);
                                        //oDS_PH_PY115A.SetValue("U_DEDAT1", 0, "20121231");
                                    }
                                }

                                //정산종료일자
                                if (oMast.TermDate == "")
                                {
                                    oDS_PH_PY115A.SetValue("U_ENDRET", 0, DateTime.Now.ToString("yyyyMMdd"));

                                    //2012년12월31일 이전 기산일, 퇴사일
                                    if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()) >= 20130101)
                                    {
                                        oDS_PH_PY115A.SetValue("U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim());
                                    }
                                    else
                                    {
                                        //oDS_PH_PY115A.SetValue("U_DSDAT2", 0, "20130101");
                                    }

                                    oDS_PH_PY115A.SetValue("U_DEDAT2", 0, DateTime.Now.ToString("yyyyMMdd"));

                                    oForm.Items.Item("JSNGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_Index);
                                    oDS_PH_PY115A.SetValue("U_JSNNAM", 0, oForm.Items.Item("JSNGBN").Specific.Selected.Description);
                                    oForm.Items.Item("ENDRET").Enabled = true;
                                    oForm.Items.Item("ENDRET").Click(BoCellClickType.ct_Regular);
                                    MsterChk = true;
                                }
                                else
                                {
                                    oDS_PH_PY115A.SetValue("U_ENDRET", 0, oMast.TermDate);

                                    if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()) >= 20130101)
                                    {
                                        oDS_PH_PY115A.SetValue("U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim());
                                    }
                                    else
                                    {
                                        //2012년12월31일 이전 기산일, 퇴사일
                                        oDS_PH_PY115A.SetValue("U_DSDAT2", 0, "20130101");
                                    }

                                    oDS_PH_PY115A.SetValue("U_DEDAT2", 0, oMast.TermDate);


                                    oForm.Items.Item("JSNGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oDS_PH_PY115A.SetValue("U_JSNNAM", 0, oForm.Items.Item("JSNGBN").Specific.Selected.Description);
                                    oForm.Items.Item("ENDRET").Enabled = false;
                                    MsterChk = false;
                                }

                                //임원누진율
                                Compute_APPRAT();
                            }
                        }

                        oForm.Items.Item("MSTNAM").Update();
                        oForm.Items.Item("EmpID").Update();
                        oForm.Items.Item("INPDAT").Update();
                        oForm.Items.Item("OUTDAT").Update();
                        oForm.Items.Item("STRRET").Update();
                        oForm.Items.Item("ENDRET").Update();
                        Compute_GNSDAY();
                        Compute_GNMYER();
                        break;

                    case "ENDRET":

                        MsterChk = true;

                        //oDS_PH_PY115A.setValue "U_DSDAT2", 0, Format$("2013-01-01", "yyyymmdd")
                        if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()) >= 20130101)
                        {
                            oDS_PH_PY115A.SetValue("U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim());
                        }
                        else
                        {
                            oDS_PH_PY115A.SetValue("U_DSDAT2", 0, "20130101");
                        }

                        oDS_PH_PY115A.SetValue("U_DEDAT2", 0, oDS_PH_PY115A.GetValue("U_ENDRET", 0).ToString());

                        Compute_GNSDAY();
                        break;

                    case "GNSYER":
                    case "GNSMON":
                    case "GNSDAY":

                        Compute_GNMYER();
                        break;

                    case "PAYSD1":
                    case "PAYED1":

                        if (oDS_PH_PY115A.GetValue("U_PAYED1", 0).Trim() == "")
                        {
                            oDS_PH_PY115A.SetValue("U_PAYED1", 0, codeHelpClass.Mid(oDS_PH_PY115A.GetValue("U_PAYSD1", 0).Trim(), 0, 6) + dataHelpClass.Lday(oDS_PH_PY115A.GetValue("U_PAYSD1", 0).Trim()));
                            oForm.Items.Item("PAYED1").Update();
                        }
                        oDS_PH_PY115A.SetValue("U_PAYDY1", 0, (dataHelpClass.TermDay(oDS_PH_PY115A.GetValue("U_PAYSD1", 0).Trim(), oDS_PH_PY115A.GetValue("U_PAYED1", 0).Trim())).ToString());
                        oForm.Items.Item("PAYDY1").Update();
                        break;

                    case "PAYSD2":
                    case "PAYED2":
                        {
                            if (oDS_PH_PY115A.GetValue("U_PAYED2", 0).Trim() == "")
                            {
                                oDS_PH_PY115A.SetValue("U_PAYED2", 0, codeHelpClass.Mid(oDS_PH_PY115A.GetValue("U_PAYSD2", 0).Trim(), 0, 6) + dataHelpClass.Lday(oDS_PH_PY115A.GetValue("U_PAYSD2", 0).Trim()));
                                oForm.Items.Item("PAYED2").Update();
                            }
                            oDS_PH_PY115A.SetValue("U_PAYDY2", 0, dataHelpClass.TermDay(oDS_PH_PY115A.GetValue("U_PAYSD2", 0).Trim(), oDS_PH_PY115A.GetValue("U_PAYED2", 0).Trim()).ToString());
                            oForm.Items.Item("PAYDY2").Update();
                            break;
                        }

                    case "PAYSD3":
                    case "PAYED3":

                        if (oDS_PH_PY115A.GetValue("U_PAYED3", 0).Trim() == "")
                        {
                            oDS_PH_PY115A.SetValue("U_PAYED3", 0, codeHelpClass.Mid(oDS_PH_PY115A.GetValue("U_PAYSD3", 0), 0, 6) + dataHelpClass.Lday(oDS_PH_PY115A.GetValue("U_PAYSD3", 0).Trim()));
                            oForm.Items.Item("PAYED3").Update();
                        }
                        oDS_PH_PY115A.SetValue("U_PAYDY3", 0, dataHelpClass.TermDay(oDS_PH_PY115A.GetValue("U_PAYSD3", 0).Trim(), oDS_PH_PY115A.GetValue("U_PAYED3", 0).Trim()).ToString());
                        oForm.Items.Item("PAYDY3").Update();
                        break;

                    case "PAYSD4":
                    case "PAYED4":

                        if (oDS_PH_PY115A.GetValue("U_PAYED2", 0).Trim() == "")
                        {
                            oDS_PH_PY115A.SetValue("U_PAYED4", 0, codeHelpClass.Mid(oDS_PH_PY115A.GetValue("U_PAYSD4", 0).Trim(), 0, 6) + dataHelpClass.Lday(oDS_PH_PY115A.GetValue("U_PAYSD4", 0).Trim()));
                            oForm.Items.Item("PAYED4").Update();
                        }
                        oDS_PH_PY115A.SetValue("U_PAYDY4", 0, Convert.ToString(dataHelpClass.TermDay(oDS_PH_PY115A.GetValue("U_PAYSD4", 0).Trim(), oDS_PH_PY115A.GetValue("U_PAYED4", 0).Trim())));
                        oForm.Items.Item("PAYDY4").Update();
                        break;

                    case "JOTDAT":
                    case "JINDAT":

                        STRDAT = oDS_PH_PY115A.GetValue("U_JINDAT", 0).Trim();
                        ENDDAT = oDS_PH_PY115A.GetValue("U_JOTDAT", 0).Trim();

                        if (STRDAT != "" && ENDDAT != "")
                        {
                            dataHelpClass.Term2(STRDAT, ENDDAT);
                            //근무년월일
                            WK_JONYER = PSH_Globals.ZPAY_GBL_GNSYER;
                            WK_JONMON = PSH_Globals.ZPAY_GBL_GNSMON;
                            if (PSH_Globals.ZPAY_GBL_GNSDAY > 0)
                            {
                                WK_JONMON = Convert.ToInt16(WK_JONMON + 1);
                            }

                            WK_GNMDAY = Convert.ToInt16(WK_JONYER * 12 + WK_JONMON);
                            //중복월수
                            if (Convert.ToDouble(STRDAT) < Convert.ToDouble(oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim()))
                            {
                                STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0).Trim();
                            }

                            if (Convert.ToDouble(ENDDAT) > Convert.ToDouble(oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim()))
                            {
                                ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0).Trim();
                            }

                            dataHelpClass.Term2(STRDAT, ENDDAT);
                            WK_DUPMON = Convert.ToInt16(PSH_Globals.ZPAY_GBL_GNSYER * 12);

                            if (PSH_Globals.ZPAY_GBL_GNSDAY > 0)
                            {
                                WK_DUPMON = Convert.ToInt16(WK_DUPMON + PSH_Globals.ZPAY_GBL_GNSMON + 1);
                            }
                            else
                            {
                                WK_DUPMON = Convert.ToInt16(WK_DUPMON + PSH_Globals.ZPAY_GBL_GNSMON);
                            }

                            oDS_PH_PY115A.SetValue("U_JONYER", 0, WK_JONYER.ToString()); //종전근속년수
                            oDS_PH_PY115A.SetValue("U_JONMON", 0, WK_JONMON.ToString()); //종전근속월수
                            oDS_PH_PY115A.SetValue("U_GNMDAY", 0, WK_GNMDAY.ToString()); //종전세법상근무년수
                            oDS_PH_PY115A.SetValue("U_JMMMON", 0, WK_DUPMON.ToString()); //중복월수
                        }

                        break;

                    case "JRET01":
                    case "JYIL03":

                        JRET01 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JRET01", 0).Trim());
                        JYIL03 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYIL03", 0).Trim());
                        JTOT01 = JRET01 + JYIL03;
                        oDS_PH_PY115A.SetValue("U_JTOT01", 0, JTOT01.ToString());
                        oForm.Items.Item("JTOT01").Update();
                        break;

                    case "JADD01":
                    case "JYIL04":

                        JADD01 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JADD01", 0).Trim());
                        JYIL04 = Convert.ToDouble(oDS_PH_PY115A.GetValue("U_JYIL04", 0).Trim());
                        JTOT02 = JADD01 + JYIL04;
                        oDS_PH_PY115A.SetValue("U_JTOT02", 0, JTOT02.ToString());
                        oForm.Items.Item("JTOT02").Update();
                        break;

                    case "JIGBIL":

                        STRDAT = oDS_PH_PY115A.GetValue("U_JIGBIL", 0).Trim();
                        //신고월 설정하기
                        if (STRDAT != "")
                        {
                            if (Convert.ToInt16(codeHelpClass.Mid(STRDAT, 6, 2)) <= 10)
                            {
                                oDS_PH_PY115A.SetValue("U_SINYMM", 0, codeHelpClass.Left(STRDAT, 6));
                            }
                            else
                            {
                                oDS_PH_PY115A.SetValue("U_SINYMM", 0, Convert.ToDateTime(dataHelpClass.ConvertDateType(STRDAT, "-")).AddMonths(1).ToString("yyyyMM"));
                            }

                            //switch (codeHelpClass.Mid(STRDAT, 6, 2))
                            //{
                            //    case object _ when Strings.Mid(STRDAT, 7, 2) <= System.Convert.ToString(10):

                            //        oDS_PH_PY115A.setValue("U_SINYMM", 0, Strings.Left(STRDAT, 6));
                            //        break;

                            //    default:

                            //        oDS_PH_PY115A.setValue("U_SINYMM", 0, VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, (DateTime)VB6.Format(STRDAT, "0000-00-00")), "YYYYMM"));
                            //        break;

                            //}
                            oForm.Items.Item("SINYMM").Update();
                        }

                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) //추가및 업데이트시에
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (dataHelpClass.Value_ChkYn("[@PH_PY115A]", "U_MSTCOD", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", " AND U_ENDRET = '" + oForm.Items.Item("ENDRET").Specific.Value + "'") == false)
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            if (PH_PY115_DataValidCheck("") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else if (Convert.ToDouble(oDS_PH_PY115A.GetValue("U_RETPAY", 0).Trim()) <= 0)
                            {
                                BubbleEvent = false;
                                PSH_Globals.SBO_Application.StatusBar.SetText("퇴직급여의 값이 0입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                            else if (MsterChk == true)
                            {
                                MasterUpdate();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CBtn1" && oForm.Items.Item("MSTCOD").Enabled == true)
                    {
                        oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                    }
                    else if (pVal.ItemUID == "Btn3" && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) //세액 계산
                    {
                        if (Compute_TAX() == true && Compute_TAX_3() == true && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                    else if (pVal.ItemUID == "Btn2" && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) //퇴직금 계산
                    {
                        if (PH_PY115_DataValidCheck(pVal.ItemUID) == false)
                        {
                            BubbleEvent = false;
                        }
                        else if (Compute_PAY() == true && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                    else if (pVal.ItemUID == "Btn1" && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) //퇴직금 생성
                    {
                        if (PH_PY115_DataValidCheck("") == false)
                        {
                            BubbleEvent = false;
                        }
                        else if (Create_Data() == true && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                    else if (pVal.ItemUID == "LBtn1")
                    {
                        //if (System.Convert.ToBoolean(oDS_PH_PY115A.GetValue("U_MSTCOD", 0)))
                        //{
                        //}

                        PSH_BaseClass oTmpObject = new PH_PY001();

                        if (oTmpObject != null)
                        {
                            oTmpObject.LoadForm("");

                            //if (PSH_Globals.SBO_Application.Forms.ActiveForm.Type == "PH_PY001")
                            //{
                            PSH_Globals.SBO_Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PSH_Globals.SBO_Application.Forms.ActiveForm.Freeze(true);
                            PSH_Globals.SBO_Application.Forms.ActiveForm.Items.Item("Code").Specific.Value = oDS_PH_PY115A.GetValue("U_MSTCOD", 0);
                            PSH_Globals.SBO_Application.Forms.ActiveForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.Forms.ActiveForm.Freeze(false);

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpObject);

                            BubbleEvent = false;
                            return;
                            //}
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ActionSuccess == true && (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                    {
                        PH_PY115_FormItemEnabled();
                    }

                    if (pVal.ItemUID == "Btn3") //환산급여별공제 계산을 위하여 Compute_TAX 재 실행
                    {
                        Compute_TAX();
                        Compute_TAX_3();
                    }

                    if (pVal.ItemUID == "Folder1")
                    {
                        oForm.PaneLevel = 1;
                    }
                    else if (pVal.ItemUID == "Folder2")
                    {
                        oForm.PaneLevel = 2;
                    }
                    else if (pVal.ItemUID == "Folder3")
                    {
                        oForm.PaneLevel = 3;
                    }
                    else if (pVal.ItemUID == "Folder4")
                    {
                        oForm.PaneLevel = 4;
                    }
                    else if (pVal.ItemUID == "Folder5")
                    {
                        oForm.PaneLevel = 5;
                    }
                    else if (pVal.ItemUID == "Folder6")
                    {
                        oForm.PaneLevel = 6;
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "MSTCOD" && pVal.CharPressed == 9 && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oForm.Items.Item(pVal.ItemUID).Specific.String + "'", "") == true)
                        {
                            oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "J01NBR" && pVal.CharPressed == 9 && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && oForm.Items.Item("J01NAM").Specific.Value.ToString().Trim() != "")
                    {
                        if (oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() == "")
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("종전 사업자번호를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                        else //사업자번호 체크
                        {
                            if (oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim().Length <= 12)
                            {
                                if (dataHelpClass.TaxNoCheck(oForm.Items.Item(pVal.ItemUID).Specific.String) == false)
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("사업자번호가 틀립니다. 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
                        case "Grid1":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }

                            break;

                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID != "1000001" && pVal.ItemUID != "2" && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        if (oLastItemUID == "MSTCOD")
                        {
                            if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oForm.Items.Item(oLastItemUID).Specific.Value.ToString().Trim() + "'", "") == true && oForm.Items.Item(oLastItemUID).Specific.Value.ToString().Trim() != "" && oLastItemUID != pVal.ItemUID)
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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            int i;
            int j;
            int[] PAYTOT = new int[4];

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "MSTCOD")
                        {
                            FlushToItemValue(pVal.ItemUID, 0);
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "PAYSD1" || pVal.ItemUID == "PAYSD2" || pVal.ItemUID == "PAYSD3" || pVal.ItemUID == "PAYSD4" || pVal.ItemUID == "PAYED1" || pVal.ItemUID == "PAYED2" || pVal.ItemUID == "PAYED3" || pVal.ItemUID == "PAYED4" || pVal.ItemUID == "JINDAT" || pVal.ItemUID == "JOTDAT" || pVal.ItemUID == "JIGBIL" || pVal.ItemUID == "JRET01" || pVal.ItemUID == "JADD01" || pVal.ItemUID == "JYIL03" || pVal.ItemUID == "JYIL04"))
                        {
                            FlushToItemValue(pVal.ItemUID, 0);
                        }
                           else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && (pVal.ItemUID == "GNSYER" || pVal.ItemUID == "GNSMON" || pVal.ItemUID == "GNSDAY" || pVal.ItemUID == "ENDRET"))
                        {
                            FlushToItemValue(pVal.ItemUID, 0);
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "PAYT01" || pVal.ItemUID == "PAYT02" || pVal.ItemUID == "PAYT03" || pVal.ItemUID == "PAYT04" || pVal.ItemUID == "BNST01" || pVal.ItemUID == "BNST02" || pVal.ItemUID == "BNST03" || pVal.ItemUID == "BNST04" || pVal.ItemUID == "BNST05" || pVal.ItemUID == "BNST06" || pVal.ItemUID == "BNST07" || pVal.ItemUID == "BNST08" || pVal.ItemUID == "BNST09" || pVal.ItemUID == "BNST10" || pVal.ItemUID == "BNST11" || pVal.ItemUID == "BNST12" || pVal.ItemUID == "YONCSU" || pVal.ItemUID == "SUDAMT" || pVal.ItemUID == "GITAA2" || pVal.ItemUID == "ExPay2" || pVal.ItemUID == "GITAA1" || pVal.ItemUID == "TGSDAY" || pVal.ItemUID == "ADDRAT"))
                        {
                            if (Compute_PAY() == true)
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                }
                            }
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "JYNTOT" || pVal.ItemUID == "JYNWON" || pVal.ItemUID == "JYNBUL" || pVal.ItemUID == "JYNGON" || pVal.ItemUID == "MYNTOT" || pVal.ItemUID == "MYNWON" || pVal.ItemUID == "MYNBUL" || pVal.ItemUID == "MYNGON"))
                        {
                            Display_ilSiGum();
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "SH1JIG" || pVal.ItemUID == "SH3JIG"))
                        {
                            Display_TaxWhanSan(1, oDS_PH_PY115A.GetValue("U_BUBCHK", 0).Trim());
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "SH2JIG" || pVal.ItemUID == "SH4JIG"))
                        {
                            Display_TaxWhanSan(2, oDS_PH_PY115A.GetValue("U_BUBCHK", 0).Trim());
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && (pVal.ItemUID == "GULJUM" || pVal.ItemUID == "JONGAB" || pVal.ItemUID == "JONJUM" || pVal.ItemUID == "GITJA1" || pVal.ItemUID == "JSNGAB" || pVal.ItemUID == "JSNJUM" || pVal.ItemUID == "MEDJSN" || pVal.ItemUID == "GITJA2"))
                        {
                            Compute_TAX_3();
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && pVal.ItemUID == "Mat1" && codeHelpClass.Left(pVal.ColUID, 6) == "CSUAMT")
                        {
                            oMat1.FlushToDataSource();
                            for (i = 0; i < PAYTOT.Length; i++)
                            {
                                PAYTOT[i] = 0;
                            }

                            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                for (j = 0; j <= 3; j++)
                                {
                                    PAYTOT[j] = PAYTOT[j] + Convert.ToInt32(Convert.ToDouble(oDS_PH_PY115B.GetValue("U_CSUAMT" + (j + 1).ToString().Trim(), i).Trim()));

                                }
                            }

                            for (i = 0; i < PAYTOT.Length; i++)
                            {
                                oDS_PH_PY115A.SetValue("U_PAYT0" + (i + 1).ToString().Trim(), 0, PAYTOT[i].ToString());
                            }

                            oForm.Items.Item("PAYT01").Update();
                            oForm.Items.Item("PAYT02").Update();
                            oForm.Items.Item("PAYT03").Update();
                            oForm.Items.Item("PAYT04").Update();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY115A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
            string Code;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ZPAY_g_EmpID oMast; //= new ZPAY_g_EmpID();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            Code = oForm.Items.Item("Code").Specific.Value;
                            sQry = "Select ENDRET = Convert(Char(8),DateAdd(dd, 1,U_ENDRET),112) From [@PH_PY115A] Where Code = '" + Code + "'";
                            oRecordSet.DoQuery(sQry);

                            oMast = dataHelpClass.Get_EmpID_InFo(oForm.Items.Item("MSTCOD").Specific.Value);

                            if (oRecordSet.Fields.Item(0).Value > oMast.RETDAT)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                                {
                                    BubbleEvent = false;
                                    oForm.Freeze(false);
                                    return;
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("퇴직(중간)정산 확정자료는 삭제할 수 없습니다.");
                                BubbleEvent = false;
                                oForm.Freeze(false);
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY115_FormItemEnabled();
                            PH_PY115_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY115_FormItemEnabled();
                            PH_PY115_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY115_FormItemEnabled();
                            PH_PY115_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY115_FormItemEnabled();
                            break;
                        case "1293": //행삭제
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
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
