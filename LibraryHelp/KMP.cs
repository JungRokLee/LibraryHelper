using System;
using System.Collections.Generic;


namespace LibraryHelp
{
    public class KMP
    {
        public  int[] getPi(String pattern)
        {  // 접두사와 접미사 매칭 Count
            int LenOfPattern = pattern.Length;     // 찾을 Pattern's Length
            int[] pi = new int[LenOfPattern];        // pi[] 할당
            int j = 0;                               // 패턴을 탐색할 인덱스

            for (int i = 1; i < LenOfPattern; i++)
            {      // i = 1부터 
                while (j > 0 && pattern[i] != pattern[j])
                { // KMP알고리즘, 매칭되는 것 건너뛴다
                    j = pi[j - 1];                                         // j를 재할당 (건너 뜀)
                } if (pattern[i] == pattern[j])
                {            // 접두사랑 접미사가 같다면
                    pi[i] = ++j;                                         // 해당 인덱스 카운트  (길이)
                }
            }
            return pi;     // pi[] 반환
        }

        public  List<int> kmpAlgorithm(String str, String pattern)
        {  // 매칭되는 str을 ArrayList에 저장

            List<int> list = new List<int>();      // 매칭된 str을 저장할 ArrayList
            int[] pi = getPi(pattern);                       // 찾을 패턴의 접두사와 접미사가 카운트된 배열 불러온다
            int LenOfStr = str.Length;                      // 원본 str의 길이
            int LenOfPattern = pattern.Length;              // 찾을 pattern의 길이
            int j = 0;                                        // pattern을 탐색시킬 인덱스

            for (int i = 0; i < LenOfStr; i++)
            {                   // str의 0번째 (가장 처음부터)  탐색
                while (j > 0 && str[i] != pattern[j])
                {    // KMP 알고리즘의 핵심
                    j = pi[j - 1];                                       // j 재할당 (건너 뛴다)
                } if (str[i] == pattern[j])
                {               // 문자가 같을 때
                    if (j == LenOfPattern - 1)
                    {                           // 해당 인덱스가 패턴의 끝과 일치? 그럼 매칭 성공!
                        list.Add(i - LenOfPattern + 1);                    // 찾은 패턴을 list에 순서대로 넣는다
                        j = pi[j];                                     // j 초기화
                    }
                    else
                    {                                   // 해당 인덱스가 패턴의 끝과 일치하지 않으면 더 찾아야 함!
                        j++;                                  // 탐색 인덱스 증가!
                    }
                }
            }
            return list;                                       // 매칭된 패턴이 저장된 list 반환
        }
    }
}
