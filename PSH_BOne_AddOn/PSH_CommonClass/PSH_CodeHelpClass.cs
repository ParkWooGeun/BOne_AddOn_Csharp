namespace PSH_BOne_AddOn.Code
{
    /// <summary>
    /// VB6.0와 C#의 기능 차이 보완용 Class
    /// </summary>
    public class PSH_CodeHelpClass
    {
        ///// <summary>
        ///// 호출된 메소드명 리턴
        ///// </summary>
        ///// <returns>메소드명</returns>
        //public string GetCurrentMethodName()
        //{
        //    return System.Reflection.MethodBase.GetCurrentMethod().Name;
        //}

        /// <summary>
        /// Visual Basic의 Mid 함수를 Method로 구현
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pFromInt">부터</param>
        /// <param name="pToInt">꺄지</param>
        /// <returns>처리된 문자열</returns>
        public string Mid(string pString, int pFromInt, int pToInt)
        {
            if (pFromInt < pString.Length || pToInt < pString.Length)
            {
                return pString.Substring(pFromInt, pToInt);
            }
            else
            {
                return pString;
            }
        }

        /// <summary>
        /// Visual Basic의 Left 함수를 Method로 구현
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">왼쪽에서 몇번째</param>
        /// <returns>처리된 문자열</returns>
        public string Left(string pString, int pLength)
        {
            if (pString.Length < pLength)
            {
                pLength = pString.Length;
            }

            return pString.Substring(0, pLength);
        }

        /// <summary>
        /// Visual Basic의 Right 함수를 Method로 구현
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">오른쪽에서 몇번째</param>
        /// <returns>처리된 문자열</returns>
        public string Right(string pString, int pLength)
        {
            if (pString.Length < pLength)
            {
                pLength = pString.Length;
            }

            return pString.Substring(pString.Length - pLength, pLength);
        }

        /// <summary>
        /// 고정길이 문자열 반환 #1
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">길이</param>
        /// <returns></returns>
        public string GetFixedLengthString(string pString, int pLength)
        {
            if (string.IsNullOrEmpty(pString))
            {
                return new string(' ', pLength);
            }
            else if (pString.Length > pLength)
            {
                return pString.Substring(0, pLength);
            }
            else
            {
                return pString.PadRight(pLength);
            }
        }

        /// <summary>
        /// 고정길이 문자열 반환 #2(왼쪽 특정문자 채우기)
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">길이</param>
        /// <param name="pChar">빈칸을 채울 문자</param>
        /// <returns>결과값</returns>
        public string GetFixedLengthString(string pString, int pLength, char pChar)
        {
            if (string.IsNullOrEmpty(pString))
            {
                return new string(' ', pLength);
            }
            else if (pString.Length > pLength)
            {
                return pString.Substring(0, pLength);
            }
            else
            {
                return pString.PadLeft(pLength, pChar);
            }
        }

        /// <summary>
        /// 고정길이 문자열 반환 #3(오른쪽 특정문자 채우기)
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pChar">빈칸을 채울 문자</param>
        /// <param name="pLength">길이</param>
        /// <returns>결과값</returns>
        public string GetFixedLengthString(string pString, char pChar, int pLength)
        {
            if (string.IsNullOrEmpty(pString))
            {
                return new string(' ', pLength);
            }
            else if (pString.Length > pLength)
            {
                return pString.Substring(0, pLength);
            }
            else
            {
                return pString.PadRight(pLength, pChar);
            }
        }

        /// <summary>
        /// 고정길이 문자열 반환 #1(Byte 단위)
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">길이</param>
        /// <returns></returns>
        public string GetFixedLengthStringByte(string pString, int pLength)
        {
            int byteCount = System.Text.Encoding.Default.GetByteCount(pString);

            if (byteCount > pString.Length) //일반길이보다 Byte길이가 크면(즉 2Byte문자열이면)
            {
                if (string.IsNullOrEmpty(pString))
                {
                    return new string(' ', pLength - (byteCount - pString.Length));
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength - (byteCount - pString.Length));
                }
                else
                {
                    return pString.PadRight(pLength - (byteCount - pString.Length));
                }
            }
            else
            {
                if (string.IsNullOrEmpty(pString))
                {
                    return new string(' ', pLength);
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength);
                }
                else
                {
                    return pString.PadRight(pLength);
                }
            }
        }

        /// <summary>
        /// 고정길이 문자열 반환 #2(Byte 단위)(왼쪽 특정문자 채우기)
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pLength">길이</param>
        /// <param name="pChar">빈칸을 채울 문자</param>
        /// <returns>결과값</returns>
        public string GetFixedLengthStringByte(string pString, int pLength, char pChar)
        {
            int byteCount = System.Text.Encoding.Default.GetByteCount(pString);

            if (byteCount > pString.Length) //일반길이보다 Byte길이가 크면(즉 2Byte문자열이면)
            {
                if (string.IsNullOrEmpty(pString)) 
                {
                    return new string(' ', pLength - (byteCount - pString.Length));
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength - (byteCount - pString.Length));
                }
                else
                {
                    return pString.PadLeft(pLength - (byteCount - pString.Length), pChar);
                }
            }
            else
            {
                if (string.IsNullOrEmpty(pString))
                {
                    return new string(' ', pLength);
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength);
                }
                else
                {
                    return pString.PadLeft(pLength, pChar);
                }
            }
        }

        /// <summary>
        /// 고정길이 문자열 반환 #3(Byte 단위)(오른쪽 특정문자 채우기)
        /// </summary>
        /// <param name="pString">적용할 문자열</param>
        /// <param name="pChar">빈칸을 채울 문자</param>
        /// <param name="pLength">길이</param>
        /// <returns>결과값</returns>
        public string GetFixedLengthStringByte(string pString, char pChar, int pLength)
        {
            int byteCount = System.Text.Encoding.Default.GetByteCount(pString);

            if (byteCount > pString.Length) //일반길이보다 Byte길이가 크면(즉 2Byte문자열이면)
            {
                if (string.IsNullOrEmpty(pString))
                {
                    return new string(' ', pLength - (byteCount - pString.Length));
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength - (byteCount - pString.Length));
                }
                else
                {
                    return pString.PadLeft(pLength - (byteCount - pString.Length), pChar);
                }
            }
            else
            {
                if (string.IsNullOrEmpty(pString))
                {
                    return new string(' ', pLength);
                }
                else if (pString.Length > pLength)
                {
                    return pString.Substring(0, pLength);
                }
                else
                {
                    return pString.PadRight(pLength, pChar);
                }
            }
        }
    }
}
