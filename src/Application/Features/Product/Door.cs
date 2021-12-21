namespace RoyalExcelLibrary.Application.Features.Product {
    
    internal class Door : IProduct {
        public string Name => "MDF Door";
        
        public int Qty { get; set; }

        public double Height { get; set; }

        public double Width { get; set; }

        public double TopRail { get; set; }

        public double BotRail { get; set; }

        public double LeftStile { get; set; }

        public double RightStile { get; set; }

        public string FramingBead { get; set; }

        public string Edge { get; set; }

        public string PanelDetail { get; set; }

        public string Finish { get; set; }

        public string Color { get; set; }

        public double PanelDrop { get; set; }

        public string Material { get; set; }

    }

}