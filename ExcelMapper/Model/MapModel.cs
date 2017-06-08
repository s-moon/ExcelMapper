﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Model
{

    public class MapModel
    {
        public OneToMany[] OneToMany { get; set; }
        public ManyToOne[] ManyToOne { get; set; }
    }

    public class OneToMany
    {
        public string SrcFolder { get; set; }
        public string SrcFilemask { get; set; }
        public string DstFolder { get; set; }
        public string DstFileMask { get; set; }
        public string DstFileMaskValue { get; set; }
        public string StartRow { get; set; }
        public Cell[] Cells { get; set; }
    }

    public class Cell
    {
        public string SrcSheet { get; set; }
        public string SrcCell { get; set; }
        public string DstSheet { get; set; }
        public string DstCell { get; set; }
    }

    public class ManyToOne
    {
        public string SrcFolder { get; set; }
        public string SrcFilemask { get; set; }
        public string DstFolder { get; set; }
        public string DstFileMask { get; set; }
        public string DstFileMaskValue { get; set; }
        public string StartRow { get; set; }
        public Cell[] Cells { get; set; }
    }
}