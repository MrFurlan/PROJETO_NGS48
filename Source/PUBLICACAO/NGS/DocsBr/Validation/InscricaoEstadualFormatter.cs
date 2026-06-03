using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocsBr.Validation
{
    public class InscricaoEstadualFormatter
    {
        private static readonly Dictionary<string, Func<string, string>> Formatters = new Dictionary<string, Func<string, string>>()
    {
        { "AC", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}/{ie.Substring(8, 3)}-{ie.Substring(11, 2)}" },
        { "AL", ie => $"{ie}" }, 
        { "AP", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}" },
        { "AM", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}" },
        { "BA", ie => ie.Length == 8 ? $"{ie.Substring(0, 6)}-{ie.Substring(6, 2)}" : $"{ie.Substring(0, 7)}-{ie.Substring(7, 2)}" },
        { "CE", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "DF", ie => $"{ie.Substring(0, 2)} {ie.Substring(2, 6)} {ie.Substring(8, 3)}-{ie.Substring(11, 2)}" },
        { "ES", ie => $"{ie.Substring(0, 3)}.{ie.Substring(3, 3)}.{ie.Substring(6, 2)}-{ie.Substring(8, 1)}" },
        { "GO", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}" },
        { "MA", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}" },
        { "MT", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}" },
        { "MS", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "MG", ie => $"{ie.Substring(0, 3)}.{ie.Substring(3, 3)}.{ie.Substring(6, 3)}/{ie.Substring(9, 4)}" },
        { "PA", ie => $"{ie.Substring(0, 2)}-{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "PB", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "PR", ie => $"{ie.Substring(0, 8)}-{ie.Substring(8, 2)}" },
        { "PE", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 1)}.{ie.Substring(3, 3)}.{ie.Substring(6, 4)}-{ie.Substring(10, 2)}" },
        { "PI", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "RJ", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 2)}-{ie.Substring(7, 1)}" },
        { "RN", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}" },
        { "RS", ie => $"{ie.Substring(0, 3)}/{ie.Substring(3, 7)}" },
        { "RO", ie => $"{ie.Substring(0, 3)}.{ie.Substring(3, 5)}-{ie.Substring(8, 2)}" },
        { "RR", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        { "SC", ie => $"{ie.Substring(0, 3)}.{ie.Substring(3, 3)}.{ie.Substring(6, 3)}" },
        { "SP", ie => ie.Length == 12 ? $"{ie.Substring(0, 3)}.{ie.Substring(3, 3)}.{ie.Substring(6, 3)}.{ie.Substring(9, 3)}" : $"P-{ie.Substring(1, 8)}.{ie.Substring(9, 1)}/{ie.Substring(10, 3)}" },
        { "SE", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 6)}-{ie.Substring(8, 1)}" },
        /* esse trecho em Tocantins estava dando erro */
        /* { "TO", ie => $"{ie.Substring(0, 2)}.{ie.Substring(2, 2)}.{ie.Substring(4, 6)}-{ie.Substring(10, 1)}" } */
        { "TO", ie => ie.Length == 9
            ? $"{ie.Substring(0, 2)}.{ie.Substring(2, 3)}.{ie.Substring(5, 3)}-{ie.Substring(8, 1)}"
            : $"{ie.Substring(0, 2)}.{ie.Substring(2, 2)}.{ie.Substring(4, 5)}-{ie.Substring(9, 1)}" }
    };

        public static string FormatarIE(string estado, string ie)
        {
            if (Formatters.ContainsKey(estado))
            {
                try
                {
                    return Formatters[estado](ie);
                }
                catch (Exception ex)
                {
                    string t = ex.Message;
                    return "IE";
                }
            }
            return "Estado inválido ou não suportado";
        }
    }


}
