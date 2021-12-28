namespace RoyalExcelLibrary.Application.Features.Options.Materials {

    public class Material {

        public int Id { get; set; }

        public MaterialType Type { get; set; }

        public double Dimension { get; set; }

        public decimal Price { get; set; }

    }

    public class MaterialType {

        public int TypeId { get; set; }

        public string MaterialName { get; set; }

        public string CutListCode { get; set; }

    }

}
