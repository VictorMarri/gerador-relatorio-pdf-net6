﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeradorDeRelatoriosEmPDF
{
    [Serializable]
    public class Profissao
    {
        public int IdProfissao { get; set; }
        public string Nome { get; set; }
    }
}
