using RoyalExcelLibrary.Application.Common;
using RoyalExcelLibrary.Application.Features.Configuration;
using RoyalExcelLibrary.Application.Features.Options.Materials;
using System;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Application.Features.Product {

    public class DrawerBox : ICompositeProduct<DrawerBoxPart> {

        internal DrawerBox(int qty, double height, double width, double depth, MaterialType boxMaterial, MaterialType bottomMaterial, Dictionary<string, string> extras) {
            Qty = qty;
            Height = height;
            Width = width;
            Depth = depth;
            BoxMaterial = boxMaterial;
            BottomMaterial = bottomMaterial;
            Extras = extras;
        }

        public string Name => "Drawer Box";
        public int Id { get; set; }
        public int Qty { get; private set; }
        public double Height { get; private set; }
        public double Width { get; private set; }
        public double Depth { get; private set; }
        public MaterialType BoxMaterial { get; private set; }
        public MaterialType BottomMaterial { get; private set; }

        // Extra options for the box
        public IReadOnlyDictionary<string, string> Extras { get; private set; }

        public Func<DrawerBox, IList<DrawerBoxPart>> PartOutStrategy { private get; set; } = GetDefaultParts;

        public IList<DrawerBoxPart> GetParts() {
            return PartOutStrategy(this);
        }

        public decimal Price(ProductOptions options) {

            decimal optionPrices = 0;
            foreach (string category in Extras.Keys) {
                string option = Extras[category];
                if (options.ContainsOption(category, option) ) {
                    optionPrices += options[category, option];
                }
            }

            return optionPrices;
        }

        private static List<DrawerBoxPart> GetDefaultParts(DrawerBox drawerbox) {

            var parts = new List<DrawerBoxPart> {
                new DrawerBoxPart(
                    name:   "Front/Back",
                    qty:    2 * drawerbox.Qty,
                    width:  drawerbox.Height,
                    length: drawerbox.Width + ManufacturingConstants.FrontBackAdj,
                    matType: drawerbox.BoxMaterial
                ),

                new DrawerBoxPart(
                    name:   "Sides",
                    qty:    2 * drawerbox.Qty,
                    width:  drawerbox.Height,
                    length: drawerbox.Depth - ManufacturingConstants.SideAdj,
                    matType: drawerbox.BoxMaterial
                ),

                new DrawerBoxPart(
                    name:   "Bottom",
                    qty:    drawerbox.Qty,
                    width:  drawerbox.Width - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                    length: drawerbox.Depth - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                    matType: drawerbox.BottomMaterial
                )
            };

            return parts;

        }

    }

    public class DrawerBoxPart : IProduct {
        public int Id { get; set; }
        public MaterialType MatType { get; private set; }

        public double Width { get; private set; }

        public double Length { get; private set; }

        public string Name { get;  private set; }

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

    /// <summary>
    /// Using the DrawerBoxBuilder allows for the creation of a drawer box step by step, and without the use of a unwieldy constructor while leaving the DrawerBox immutable
    /// </summary>
    public class DrawerBoxBuilder {

        private int _qty = 0;
        private double _width = 0;
        private double _height = 0;
        private double _depth = 0;
        private MaterialType _boxMaterial;
        private MaterialType _bottomMaterial;
        private readonly Dictionary<string, string> _extras = new Dictionary<string, string>();

        public DrawerBoxBuilder WithQty(int qty) {
            _qty = qty;
            return this;
        }

        public DrawerBoxBuilder WithWidth(double width) {
            _width = width;
            return this;
        }

        public DrawerBoxBuilder WithHeight(double height) {
            _height = height;
            return this;
        }

        public DrawerBoxBuilder WithDepth(double depth) {
            _depth = depth;
            return this;
        }

        public DrawerBoxBuilder WithBoxMaterial(MaterialType boxMaterial) {
            _boxMaterial = boxMaterial;
            return this;
        }

        public DrawerBoxBuilder WithBotMaterial(MaterialType bottomMaterial) {
            _bottomMaterial = bottomMaterial;
            return this;
        }

        public DrawerBoxBuilder WithExtra(string key, string value) {
            _extras.Add(key, value);
            return this;
        }
        public DrawerBox Build() {

            if (_qty < 1)                                   throw new InvalidOperationException("Quantity must be greater than 1");
            if (_width < 1 || _height < 1 || _depth < 1)    throw new InvalidOperationException("Dimensions must be greater than 1");
            if (_boxMaterial is null)                       throw new InvalidOperationException("Drawer box material must not be null");
            if (_bottomMaterial is null)                    throw new InvalidOperationException("Drawer box material must not be null");

            return new DrawerBox(_qty,
                                _height,
                                _width,
                                _depth,
                                _boxMaterial,
                                _bottomMaterial,
                                _extras);
        }

    }

}