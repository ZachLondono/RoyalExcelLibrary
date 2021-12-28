using System;

namespace RoyalExcelLibrary.Application.Features.Product {
    
    public class Door : IProduct {

        public int Id { get; set; }
        internal Door(int qty, double width, double height, double topRail, double botRail, double leftStile, double rightStile, double panelDrop, string framingBead, string edge, string panelDetail, string finish, string color, string material) {
            Qty = qty;
            Width = width;
            Height = height;
            TopRail = topRail;
            BotRail = botRail;
            LeftStile = leftStile;
            RightStile = rightStile;
            PanelDrop = panelDrop;
            FramingBead = framingBead;
            Edge = edge;
            PanelDetail = panelDetail;
            Finish = finish;
            Color = color;
            Material = material;
        }

        public string Name => "MDF Door";
        public int Qty { get; private set; }
        public double Height { get; private set; }
        public double Width { get; private set; }
        public double TopRail { get; private set; }
        public double BotRail { get; private set; }
        public double LeftStile { get; private set; }
        public double RightStile { get; private set; }
        public string FramingBead { get; private set; }
        public string Edge { get; private set; }
        public string PanelDetail { get; private set; }
        public string Finish { get; private set; }
        public string Color { get; private set; }
        public double PanelDrop { get; private set; }
        public string Material { get; private set; }

    }

    public class Builder {

        private int _qty = 0;
        private double _width = 0;
        private double _height = 0;
        private double _topRail = 0;
        private double _botRail = 0;
        private double _leftStile = 0;
        private double _rightStile = 0;
        private double _panelDrop = 0;
        private string _framingBead = string.Empty;
        private string _edge = string.Empty;
        private string _panelDetail = string.Empty;
        private string _finish = string.Empty;
        private string _color = string.Empty;
        private string _material = string.Empty;

        public Builder WithQty(int qty) {
            _qty = qty;
            return this;
        }

        public Builder WithWidth(double width) {
            _width = width;
            return this;
        }

        public Builder WithHeight(double height) {
            _height = height;
            return this;
        }

        public Builder WithTopRail(double topRail) {
            _topRail = topRail;
            return this;
        }

        public Builder WithBotRail(double botRail) {
            _botRail = botRail;
            return this;
        }

        public Builder WithLeftStile(double leftStile) {
            _leftStile = leftStile;
            return this;
        }

        public Builder WithRightStile(double rightStile) {
            _rightStile = rightStile;
            return this;
        }

        public Builder WithPanelDrop(double panelDrop) {
            _panelDrop = panelDrop;
            return this;
        }

        public Builder WithFramingBead(string framingBead) {
            _framingBead = framingBead;
            return this;
        }

        public Builder WithEdge(string edge) {
            _edge = edge;
            return this;
        }

        public Builder WithPanelDetail(string panelDetail) {
            _panelDetail = panelDetail;
            return this;
        }

        public Builder WithFinish(string finish) {
            _finish = finish;
            return this;
        }
        public Builder WithColor(string color) {
            _color = color;
            return this;
        }
        public Builder WithMaterial(string material) {
            _material = material;
            return this;
        }

        public Door Build() {
            if (_qty < 1)                   throw new InvalidOperationException("Quantity must be greater than 1");
            if (_width < 1 || _height < 1)  throw new InvalidOperationException("Dimensions must be greater than 1");
            if (_topRail < 1 || _botRail < 1 || _leftStile < 1 || _rightStile < 1)
                                            throw new InvalidOperationException("Stiles and Rails must be greater than 1");
            if (string.IsNullOrEmpty(_material))        throw new InvalidOperationException("Material is not set");
            if (string.IsNullOrEmpty(_framingBead))     throw new InvalidOperationException("Framing Bead is not set");
            if (string.IsNullOrEmpty(_edge))            throw new InvalidOperationException("Edge is not set");
            if (string.IsNullOrEmpty(_panelDetail))     throw new InvalidOperationException("Panel Detail is not set");
            if (string.IsNullOrEmpty(_panelDetail))     throw new InvalidOperationException("Finish is not set");
            if (string.IsNullOrEmpty(_panelDetail))     throw new InvalidOperationException("Color is not set");

            return new Door(_qty,
                            _width,
                            _height,
                            _topRail,
                            _botRail,
                            _leftStile,
                            _rightStile,
                            _panelDrop,
                            _framingBead,
                            _edge,
                            _panelDetail,
                            _finish,
                            _color,
                            _material);

        }

    }

}