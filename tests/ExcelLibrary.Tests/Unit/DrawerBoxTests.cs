using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Options;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using FluentAssertions;
using NUnit.Framework;
using System.Linq;

namespace ExcelLibrary.Tests.Unit {
    internal class DrawerBoxTests {
        public object ManufacturingContstants { get; private set; }

        [Test]
        public void Should_CreateValidPartsWithDifferentMaterials([Values(MaterialType.EconomyBirch, MaterialType.SolidBirch, MaterialType.HybridBirch, MaterialType.Walnut, MaterialType.WhiteOak)] MaterialType sideMaterial,
                                                                    [Values] bool scoopFront) {

            // Arrange
            double height = 104;
            double width = 533;
            double depth = 533;
            DrawerBox box = new DrawerBox {

                Qty = 1,

                BottomMaterial = MaterialType.Plywood1_2,
                SideMaterial = sideMaterial,
                
                Height = height,
                Width = width,
                Depth = depth,

                ClipsOption = Clips.No_Clips,
                NotchOption = UndermountNotch.No_Notch,
                InsertOption = "insert",
                MountingHoles = false,
                ScoopFront = scoopFront,
                Logo = false,
                PostFinish = false

            };

            // Act
            var parts = box.GetParts();

            // Assert

            parts.Should().NotBeNullOrEmpty();
            // 4 sides and a bottom
            parts.Sum(p => p.Qty).Should().Be(5);
            var heights = parts.Where(p => p.Width == height && (p as DrawerBoxPart).PartType == DBPartType.Side);
            heights.Sum(p => p.Qty).Should().Be(4);
            // Front and back are total width
            var fronts = parts.Where(p => p.Length > width && p.Width == height && (p as DrawerBoxPart).PartType == DBPartType.Side);
            fronts.Sum(p => p.Qty).Should().Be(2);
            // Sides are less than the depth
            var sides = parts.Where(p => p.Length < depth && p.Width == height && (p as DrawerBoxPart).PartType == DBPartType.Side);
            sides.Sum(p => p.Qty).Should().Be(2);
            // Bottom should be less thand width and depth
            var bottom = parts.Where(p => p.Width < width && p.Length < depth && (p as DrawerBoxPart).PartType != DBPartType.Side && p.Material == MaterialType.Plywood1_2);
            bottom.Sum(p => p.Qty).Should().Be(1);

        }

    }
}
