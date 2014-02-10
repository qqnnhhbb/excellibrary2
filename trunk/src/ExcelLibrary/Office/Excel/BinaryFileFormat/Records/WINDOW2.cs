using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace ExcelLibrary.BinaryFileFormat
{
    /// <summary>
    /// General   settings   for   the  document  window  and  global  workbook   settings
    /// </summary>
    public partial class WINDOW2 : Record
    {
        public WINDOW2(Record record) : base(record) { }

        public WINDOW2()
        {
            this.Type = RecordType.WINDOW2;
        }

        public UInt16 OptionFlags;
        public UInt16 IndexFirstVisibleRow;
        public UInt16 IndexFirstVisibleColumn;
        public UInt32 ColorIndex;
        public UInt16 CachedMagFactorPB;
        public UInt16 CachedMagFactorNV;
        public UInt32 NotUsed2;

        public override void Decode()
        {
            MemoryStream stream = new MemoryStream(Data);
            BinaryReader reader = new BinaryReader(stream);
            this.OptionFlags = reader.ReadUInt16();
            this.IndexFirstVisibleRow = reader.ReadUInt16();
            this.IndexFirstVisibleColumn = reader.ReadUInt16();
            this.ColorIndex = reader.ReadUInt32();
            this.CachedMagFactorPB = reader.ReadUInt16();
            this.CachedMagFactorNV = reader.ReadUInt16();
            this.NotUsed2 = reader.ReadUInt32();
        }

        public override void Encode()
        {
            MemoryStream stream = new MemoryStream();
            BinaryWriter writer = new BinaryWriter(stream);

            writer.Write(OptionFlags);
            writer.Write(IndexFirstVisibleRow);
            writer.Write(IndexFirstVisibleColumn);
            writer.Write(ColorIndex);
            writer.Write(CachedMagFactorPB);
            writer.Write(CachedMagFactorNV);
            writer.Write(NotUsed2);

            this.Data = stream.ToArray();
            this.Size = (UInt16)Data.Length;
            base.Encode();
        }
    }
}
