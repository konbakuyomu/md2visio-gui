using md2visio.struc.figure;
using Microsoft.Office.Interop.Visio;

namespace md2visio.vsdx.@base
{
    internal class VBoundary
    {
        protected bool wild = false; // false: pinx/piny/width/height(0/0/0/0) is valid
        protected double pinx, piny, width, height;
        public double PinX { get { return pinx; } }
        public double PinY { get { return piny; } }
        public double Width { get { return width; } }
        public double Height { get { return height; } }
        public double Left
        {
            get { return pinx - width / 2; }
            set
            {
                double r = Right;
                if (value > r) return;
                pinx = (value + r) / 2;
                width = r - value;
            }
        }

        public double Right
        {
            get { return pinx + width / 2; }
            set
            {
                double L = Left;
                if (value < L) return;
                pinx = (L + value) / 2;
                width = value - L;
            }
        }

        public double Top
        {
            get { return piny + height / 2; }
            set
            {
                double b = Bottom;
                if (value < b) return;
                piny = (b + value) / 2;
                height = value - b;
            }
        }
        public double Bottom
        {
            get { return piny - height / 2; }
            set
            {
                double t = Top;
                if (value > t) return;
                piny = (value + t) / 2;
                height = t - value;
            }
        }

        public VBoundary() { }

        public VBoundary(bool wild = false) {
            this.wild = wild;
        }

        public VBoundary(double pinx, double piny, double width, double height)
        {
            this.pinx = pinx;
            this.piny = piny;
            this.width = width;
            this.height = height;
        }

        public void AlignLeft(double left)
        {
            pinx = left + width / 2;
        }

        public void AlignTop(double top)
        {
            piny = top - height / 2;
        }

        public void AlignLeftTop(double left, double top)
        {
            AlignLeft(left);
            AlignTop(top);
        }

        public void AlignRight(double right)
        {
            pinx = right - width / 2;
        }

        public void AlignBottom(double bottom)
        {
            piny = bottom + height / 2;
        }

        public void AlignRightBottom(double right, double bottom)
        {
            AlignRight(right);
            AlignBottom(bottom);
        }

        public VBoundary Clone()
        {
            VBoundary clone = new VBoundary();
            clone.pinx = pinx;
            clone.piny = piny;
            clone.width = width;
            clone.height = height;
            return clone;
        }

        public void Expand(VBoundary boundary)
        {
            if(wild)
            {
                pinx = boundary.pinx;   piny = boundary.piny;
                width = boundary.width; height = boundary.height;
                wild = false;
                return;
            }
            ExpandBoundary(this, boundary);
        }

        public void Expand(Shape shape)
        {
            Expand(new VShapeBoundary(shape));
        }

        public void Grow(Shape shape, GrowthDirection direct, double space)
        {
            if (direct.H != 0)
            {
                if (direct.H > 0) VShapeDrawer.AlignLeft(shape, Right + space);
                else VShapeDrawer.AlignRight(shape, Left - space);
            }
            else if (direct.V != 0)
            {
                if (direct.V > 0) VShapeDrawer.AlignBottom(shape, Top + space);
                else VShapeDrawer.AlignTop(shape, Bottom - space);
            }
        }

        public override string ToString()
        {
            return $"L:{Left} T:{Top} R:{Right} B:{Bottom}";
        }

        public static void ExpandBoundary(VBoundary bnd2expand, VBoundary bndCompare)
        {
            if (bndCompare.IsEmpty()) return;

            if (bndCompare.Left < bnd2expand.Left) bnd2expand.Left = bndCompare.Left;
            if (bndCompare.Right > bnd2expand.Right) bnd2expand.Right = bndCompare.Right;
            if (bndCompare.Top > bnd2expand.Top) bnd2expand.Top = bndCompare.Top;
            if (bndCompare.Bottom < bnd2expand.Bottom) bnd2expand.Bottom = bndCompare.Bottom;
        }

    }
}
