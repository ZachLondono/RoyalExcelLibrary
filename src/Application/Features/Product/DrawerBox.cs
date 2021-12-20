using RoyalExcelLibrary.Application.Common;
using RoyalExcelLibrary.Application.Features.Options.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Product {

    internal class DrawerBox : ICompositeProduct<DrawerBoxPart> {
        
        public string Name => "Drawer Box";
        public string Description => "Dovetail Drawer Box";

        public int Qty { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public double Depth { get; set; }
        public MaterialType BoxMaterial { get; set; }
        public MaterialType BottomMaterial { get; set; }

        public IList<DrawerBoxPart> GetParts() {

            IList<DrawerBoxPart> parts;
            switch (BoxMaterial.MaterialName) {
                case "Economy Birch":
                    parts = GetEconomyBirchParts();
                    break;
                case "Solid Birch":
                    parts = GetSolidBirchParts();
                    break;
                default:
                    parts = GetDefaultParts();
                    break;
            }

            return parts;

        }

        public decimal Price() {
            throw new NotImplementedException();
        }

        private List<DrawerBoxPart> GetDefaultParts() {

            var parts = new List<DrawerBoxPart>();

            parts.Add(new DrawerBoxPart(
                    name:   "Front/Back", 
                    qty:    2 * Qty,
                    width:  Height,
                    length: Width + ManufacturingConstants.FrontBackAdj,
                    matType: BoxMaterial
                ));

            parts.Add(new DrawerBoxPart(
                    name:   "Sides",
                    qty:    2 * Qty,
                    width:  Height,
                    length: Depth - ManufacturingConstants.SideAdj,
                    matType: BoxMaterial
                ));

            parts.Add(new DrawerBoxPart(
                    name:   "Bottom",
                    qty:    Qty,
                    width:  Width - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                    length: Depth - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                    matType: BottomMaterial
                ));

            return parts;

        }

        private List<DrawerBoxPart> GetEconomyBirchParts() {
            var parts = new List<DrawerBoxPart>();
            return parts;
        }

        private List<DrawerBoxPart> GetSolidBirchParts() {
            var parts = new List<DrawerBoxPart>();
            return parts;
        }

    }

    internal class DrawerBoxPart : IProduct {

        public MaterialType MatType { get; private set; }

        public double Width { get; private set; }

        public double Length { get; private set; }

        public string Name { get;  private set; }

        public string Description => "Drawer Box Part";

        public int Qty { get; private set;  }

        public DrawerBoxPart(string name, int qty, double width, double length, MaterialType matType) {
            Name = name;
            Qty = qty;
            Width = width;
            Length = length;
            MatType = matType;
        }

        public decimal Price() {
            return 0;
        }

    }

}