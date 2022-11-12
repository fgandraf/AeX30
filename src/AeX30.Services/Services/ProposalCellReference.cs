using System;
using AeX30.Services.ProposalVersion;

namespace AeX30.Services
{
    public abstract class ProposalCellReference
    {
        public static string[] Get(string footer)
        {
            switch (footer)
            {
                case "Vigência: 27/11/2017":
                case "&9Vigência: 27/11/2017":
                    return PFUI2017.References;

                case "Vigência: 11/02/2018":
                case "Vigência: 01/10/2018":
                    return PFUI2018a.References;


                case "Vigência: 26/02/2018":
                    return PFUI2018b.References;


                case "&9Vigência: 22/05/2018":
                    return PFUI2018c.References;


                case "Vigência: 12/04/2019":
                case "Vigência: 23/05/2019":
                case "Vigência: 24/05/2019":
                    return PFUI2019.References;


                case "Vigência: 13/07/2020":
                case "Vigência: 23/07/2020":
                    return PFUI2020a.References;

                case "Vigência: 24/09/2020":
                case "Vigência: 26/02/2021":
                    return PFUI2020b.References;

                case "Vigência: 04/12/2020":
                    return PFUI2020c.References;

                case "Vigência: 01/06/2021":
                case "Vigência: 05/07/2021":
                case "Vigência: 14/07/2021":
                case "Vigência: 06/08/2021":
                    return PCI2021a.References;

                case "Vigência: 21/10/2021":
                case "Vigência: 04/11/2021":
                case "Vigência: 28/03/2022":
                    return PCI2021b.References;

                case "Vigência: 28/06/2022":
                    return PCI2022a.References;

                default:
                    return null;
            }
        }
    }
}
