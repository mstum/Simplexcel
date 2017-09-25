namespace Simplexcel
{
    /// <summary>
    /// Specify a Color, as System.Drawing.Color isn't used.
    /// </summary>
    public struct Color
    {
        /// <summary>
        /// The Red component of the Color
        /// </summary>
        public readonly byte R;

        /// <summary>
        /// The Green component of the Color
        /// </summary>
        public readonly byte G;

        /// <summary>
        /// The Blue component of the Color
        /// </summary>
        public readonly byte B;

        /// <summary>
        /// The Alpha component of the Color (0 = Fully Transparent)
        /// </summary>
        public readonly byte A;

        private Color(byte alpha, byte red, byte green, byte blue)
        {
            A = alpha;
            R = red;
            G = green;
            B = blue;
        }

        /// <summary>
        /// Create a <see cref="Color"/> from the given values
        /// </summary>
        /// <param name="alpha"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        /// <returns></returns>
        public static Color FromArgb(byte alpha, byte red, byte green, byte blue)
        {
            return new Color(alpha, red, green, blue);
        }

        /// <summary>
        /// Output the Color as a Hex String in ARGB order, e.g., "FFFF0000"
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"{A:X2}{R:X2}{G:X2}{B:X2}";
        }

        /// <summary>
        /// AliceBlue (Hex #FFF0F8FF, ARGB: 255, 240, 248, 255)
        /// </summary>
        public static readonly Color AliceBlue = FromArgb(255, 240, 248, 255);

        /// <summary>
        /// AntiqueWhite (Hex #FFFAEBD7, ARGB: 255, 250, 235, 215)
        /// </summary>
        public static readonly Color AntiqueWhite = FromArgb(255, 250, 235, 215);

        /// <summary>
        /// Aqua (Hex #FF00FFFF, ARGB: 255, 0, 255, 255)
        /// </summary>
        public static readonly Color Aqua = FromArgb(255, 0, 255, 255);

        /// <summary>
        /// Aquamarine (Hex #FF7FFFD4, ARGB: 255, 127, 255, 212)
        /// </summary>
        public static readonly Color Aquamarine = FromArgb(255, 127, 255, 212);

        /// <summary>
        /// Azure (Hex #FFF0FFFF, ARGB: 255, 240, 255, 255)
        /// </summary>
        public static readonly Color Azure = FromArgb(255, 240, 255, 255);

        /// <summary>
        /// Beige (Hex #FFF5F5DC, ARGB: 255, 245, 245, 220)
        /// </summary>
        public static readonly Color Beige = FromArgb(255, 245, 245, 220);

        /// <summary>
        /// Bisque (Hex #FFFFE4C4, ARGB: 255, 255, 228, 196)
        /// </summary>
        public static readonly Color Bisque = FromArgb(255, 255, 228, 196);

        /// <summary>
        /// Black (Hex #FF000000, ARGB: 255, 0, 0, 0)
        /// </summary>
        public static readonly Color Black = FromArgb(255, 0, 0, 0);

        /// <summary>
        /// BlanchedAlmond (Hex #FFFFEBCD, ARGB: 255, 255, 235, 205)
        /// </summary>
        public static readonly Color BlanchedAlmond = FromArgb(255, 255, 235, 205);

        /// <summary>
        /// Blue (Hex #FF0000FF, ARGB: 255, 0, 0, 255)
        /// </summary>
        public static readonly Color Blue = FromArgb(255, 0, 0, 255);

        /// <summary>
        /// BlueViolet (Hex #FF8A2BE2, ARGB: 255, 138, 43, 226)
        /// </summary>
        public static readonly Color BlueViolet = FromArgb(255, 138, 43, 226);

        /// <summary>
        /// Brown (Hex #FFA52A2A, ARGB: 255, 165, 42, 42)
        /// </summary>
        public static readonly Color Brown = FromArgb(255, 165, 42, 42);

        /// <summary>
        /// BurlyWood (Hex #FFDEB887, ARGB: 255, 222, 184, 135)
        /// </summary>
        public static readonly Color BurlyWood = FromArgb(255, 222, 184, 135);

        /// <summary>
        /// CadetBlue (Hex #FF5F9EA0, ARGB: 255, 95, 158, 160)
        /// </summary>
        public static readonly Color CadetBlue = FromArgb(255, 95, 158, 160);

        /// <summary>
        /// Chartreuse (Hex #FF7FFF00, ARGB: 255, 127, 255, 0)
        /// </summary>
        public static readonly Color Chartreuse = FromArgb(255, 127, 255, 0);

        /// <summary>
        /// Chocolate (Hex #FFD2691E, ARGB: 255, 210, 105, 30)
        /// </summary>
        public static readonly Color Chocolate = FromArgb(255, 210, 105, 30);

        /// <summary>
        /// Coral (Hex #FFFF7F50, ARGB: 255, 255, 127, 80)
        /// </summary>
        public static readonly Color Coral = FromArgb(255, 255, 127, 80);

        /// <summary>
        /// CornflowerBlue (Hex #FF6495ED, ARGB: 255, 100, 149, 237)
        /// </summary>
        public static readonly Color CornflowerBlue = FromArgb(255, 100, 149, 237);

        /// <summary>
        /// Cornsilk (Hex #FFFFF8DC, ARGB: 255, 255, 248, 220)
        /// </summary>
        public static readonly Color Cornsilk = FromArgb(255, 255, 248, 220);

        /// <summary>
        /// Crimson (Hex #FFDC143C, ARGB: 255, 220, 20, 60)
        /// </summary>
        public static readonly Color Crimson = FromArgb(255, 220, 20, 60);

        /// <summary>
        /// Cyan (Hex #FF00FFFF, ARGB: 255, 0, 255, 255)
        /// </summary>
        public static readonly Color Cyan = FromArgb(255, 0, 255, 255);

        /// <summary>
        /// DarkBlue (Hex #FF00008B, ARGB: 255, 0, 0, 139)
        /// </summary>
        public static readonly Color DarkBlue = FromArgb(255, 0, 0, 139);

        /// <summary>
        /// DarkCyan (Hex #FF008B8B, ARGB: 255, 0, 139, 139)
        /// </summary>
        public static readonly Color DarkCyan = FromArgb(255, 0, 139, 139);

        /// <summary>
        /// DarkGoldenrod (Hex #FFB8860B, ARGB: 255, 184, 134, 11)
        /// </summary>
        public static readonly Color DarkGoldenrod = FromArgb(255, 184, 134, 11);

        /// <summary>
        /// DarkGray (Hex #FFA9A9A9, ARGB: 255, 169, 169, 169)
        /// </summary>
        public static readonly Color DarkGray = FromArgb(255, 169, 169, 169);

        /// <summary>
        /// DarkGreen (Hex #FF006400, ARGB: 255, 0, 100, 0)
        /// </summary>
        public static readonly Color DarkGreen = FromArgb(255, 0, 100, 0);

        /// <summary>
        /// DarkKhaki (Hex #FFBDB76B, ARGB: 255, 189, 183, 107)
        /// </summary>
        public static readonly Color DarkKhaki = FromArgb(255, 189, 183, 107);

        /// <summary>
        /// DarkMagenta (Hex #FF8B008B, ARGB: 255, 139, 0, 139)
        /// </summary>
        public static readonly Color DarkMagenta = FromArgb(255, 139, 0, 139);

        /// <summary>
        /// DarkOliveGreen (Hex #FF556B2F, ARGB: 255, 85, 107, 47)
        /// </summary>
        public static readonly Color DarkOliveGreen = FromArgb(255, 85, 107, 47);

        /// <summary>
        /// DarkOrange (Hex #FFFF8C00, ARGB: 255, 255, 140, 0)
        /// </summary>
        public static readonly Color DarkOrange = FromArgb(255, 255, 140, 0);

        /// <summary>
        /// DarkOrchid (Hex #FF9932CC, ARGB: 255, 153, 50, 204)
        /// </summary>
        public static readonly Color DarkOrchid = FromArgb(255, 153, 50, 204);

        /// <summary>
        /// DarkRed (Hex #FF8B0000, ARGB: 255, 139, 0, 0)
        /// </summary>
        public static readonly Color DarkRed = FromArgb(255, 139, 0, 0);

        /// <summary>
        /// DarkSalmon (Hex #FFE9967A, ARGB: 255, 233, 150, 122)
        /// </summary>
        public static readonly Color DarkSalmon = FromArgb(255, 233, 150, 122);

        /// <summary>
        /// DarkSeaGreen (Hex #FF8FBC8B, ARGB: 255, 143, 188, 139)
        /// </summary>
        public static readonly Color DarkSeaGreen = FromArgb(255, 143, 188, 139);

        /// <summary>
        /// DarkSlateBlue (Hex #FF483D8B, ARGB: 255, 72, 61, 139)
        /// </summary>
        public static readonly Color DarkSlateBlue = FromArgb(255, 72, 61, 139);

        /// <summary>
        /// DarkSlateGray (Hex #FF2F4F4F, ARGB: 255, 47, 79, 79)
        /// </summary>
        public static readonly Color DarkSlateGray = FromArgb(255, 47, 79, 79);

        /// <summary>
        /// DarkTurquoise (Hex #FF00CED1, ARGB: 255, 0, 206, 209)
        /// </summary>
        public static readonly Color DarkTurquoise = FromArgb(255, 0, 206, 209);

        /// <summary>
        /// DarkViolet (Hex #FF9400D3, ARGB: 255, 148, 0, 211)
        /// </summary>
        public static readonly Color DarkViolet = FromArgb(255, 148, 0, 211);

        /// <summary>
        /// DeepPink (Hex #FFFF1493, ARGB: 255, 255, 20, 147)
        /// </summary>
        public static readonly Color DeepPink = FromArgb(255, 255, 20, 147);

        /// <summary>
        /// DeepSkyBlue (Hex #FF00BFFF, ARGB: 255, 0, 191, 255)
        /// </summary>
        public static readonly Color DeepSkyBlue = FromArgb(255, 0, 191, 255);

        /// <summary>
        /// DimGray (Hex #FF696969, ARGB: 255, 105, 105, 105)
        /// </summary>
        public static readonly Color DimGray = FromArgb(255, 105, 105, 105);

        /// <summary>
        /// DodgerBlue (Hex #FF1E90FF, ARGB: 255, 30, 144, 255)
        /// </summary>
        public static readonly Color DodgerBlue = FromArgb(255, 30, 144, 255);

        /// <summary>
        /// Firebrick (Hex #FFB22222, ARGB: 255, 178, 34, 34)
        /// </summary>
        public static readonly Color Firebrick = FromArgb(255, 178, 34, 34);

        /// <summary>
        /// FloralWhite (Hex #FFFFFAF0, ARGB: 255, 255, 250, 240)
        /// </summary>
        public static readonly Color FloralWhite = FromArgb(255, 255, 250, 240);

        /// <summary>
        /// ForestGreen (Hex #FF228B22, ARGB: 255, 34, 139, 34)
        /// </summary>
        public static readonly Color ForestGreen = FromArgb(255, 34, 139, 34);

        /// <summary>
        /// Fuchsia (Hex #FFFF00FF, ARGB: 255, 255, 0, 255)
        /// </summary>
        public static readonly Color Fuchsia = FromArgb(255, 255, 0, 255);

        /// <summary>
        /// Gainsboro (Hex #FFDCDCDC, ARGB: 255, 220, 220, 220)
        /// </summary>
        public static readonly Color Gainsboro = FromArgb(255, 220, 220, 220);

        /// <summary>
        /// GhostWhite (Hex #FFF8F8FF, ARGB: 255, 248, 248, 255)
        /// </summary>
        public static readonly Color GhostWhite = FromArgb(255, 248, 248, 255);

        /// <summary>
        /// Gold (Hex #FFFFD700, ARGB: 255, 255, 215, 0)
        /// </summary>
        public static readonly Color Gold = FromArgb(255, 255, 215, 0);

        /// <summary>
        /// Goldenrod (Hex #FFDAA520, ARGB: 255, 218, 165, 32)
        /// </summary>
        public static readonly Color Goldenrod = FromArgb(255, 218, 165, 32);

        /// <summary>
        /// Gray (Hex #FF808080, ARGB: 255, 128, 128, 128)
        /// </summary>
        public static readonly Color Gray = FromArgb(255, 128, 128, 128);

        /// <summary>
        /// Green (Hex #FF008000, ARGB: 255, 0, 128, 0)
        /// </summary>
        public static readonly Color Green = FromArgb(255, 0, 128, 0);

        /// <summary>
        /// GreenYellow (Hex #FFADFF2F, ARGB: 255, 173, 255, 47)
        /// </summary>
        public static readonly Color GreenYellow = FromArgb(255, 173, 255, 47);

        /// <summary>
        /// Honeydew (Hex #FFF0FFF0, ARGB: 255, 240, 255, 240)
        /// </summary>
        public static readonly Color Honeydew = FromArgb(255, 240, 255, 240);

        /// <summary>
        /// HotPink (Hex #FFFF69B4, ARGB: 255, 255, 105, 180)
        /// </summary>
        public static readonly Color HotPink = FromArgb(255, 255, 105, 180);

        /// <summary>
        /// IndianRed (Hex #FFCD5C5C, ARGB: 255, 205, 92, 92)
        /// </summary>
        public static readonly Color IndianRed = FromArgb(255, 205, 92, 92);

        /// <summary>
        /// Indigo (Hex #FF4B0082, ARGB: 255, 75, 0, 130)
        /// </summary>
        public static readonly Color Indigo = FromArgb(255, 75, 0, 130);

        /// <summary>
        /// Ivory (Hex #FFFFFFF0, ARGB: 255, 255, 255, 240)
        /// </summary>
        public static readonly Color Ivory = FromArgb(255, 255, 255, 240);

        /// <summary>
        /// Khaki (Hex #FFF0E68C, ARGB: 255, 240, 230, 140)
        /// </summary>
        public static readonly Color Khaki = FromArgb(255, 240, 230, 140);

        /// <summary>
        /// Lavender (Hex #FFE6E6FA, ARGB: 255, 230, 230, 250)
        /// </summary>
        public static readonly Color Lavender = FromArgb(255, 230, 230, 250);

        /// <summary>
        /// LavenderBlush (Hex #FFFFF0F5, ARGB: 255, 255, 240, 245)
        /// </summary>
        public static readonly Color LavenderBlush = FromArgb(255, 255, 240, 245);

        /// <summary>
        /// LawnGreen (Hex #FF7CFC00, ARGB: 255, 124, 252, 0)
        /// </summary>
        public static readonly Color LawnGreen = FromArgb(255, 124, 252, 0);

        /// <summary>
        /// LemonChiffon (Hex #FFFFFACD, ARGB: 255, 255, 250, 205)
        /// </summary>
        public static readonly Color LemonChiffon = FromArgb(255, 255, 250, 205);

        /// <summary>
        /// LightBlue (Hex #FFADD8E6, ARGB: 255, 173, 216, 230)
        /// </summary>
        public static readonly Color LightBlue = FromArgb(255, 173, 216, 230);

        /// <summary>
        /// LightCoral (Hex #FFF08080, ARGB: 255, 240, 128, 128)
        /// </summary>
        public static readonly Color LightCoral = FromArgb(255, 240, 128, 128);

        /// <summary>
        /// LightCyan (Hex #FFE0FFFF, ARGB: 255, 224, 255, 255)
        /// </summary>
        public static readonly Color LightCyan = FromArgb(255, 224, 255, 255);

        /// <summary>
        /// LightGoldenrodYellow (Hex #FFFAFAD2, ARGB: 255, 250, 250, 210)
        /// </summary>
        public static readonly Color LightGoldenrodYellow = FromArgb(255, 250, 250, 210);

        /// <summary>
        /// LightGray (Hex #FFD3D3D3, ARGB: 255, 211, 211, 211)
        /// </summary>
        public static readonly Color LightGray = FromArgb(255, 211, 211, 211);

        /// <summary>
        /// LightGreen (Hex #FF90EE90, ARGB: 255, 144, 238, 144)
        /// </summary>
        public static readonly Color LightGreen = FromArgb(255, 144, 238, 144);

        /// <summary>
        /// LightPink (Hex #FFFFB6C1, ARGB: 255, 255, 182, 193)
        /// </summary>
        public static readonly Color LightPink = FromArgb(255, 255, 182, 193);

        /// <summary>
        /// LightSalmon (Hex #FFFFA07A, ARGB: 255, 255, 160, 122)
        /// </summary>
        public static readonly Color LightSalmon = FromArgb(255, 255, 160, 122);

        /// <summary>
        /// LightSeaGreen (Hex #FF20B2AA, ARGB: 255, 32, 178, 170)
        /// </summary>
        public static readonly Color LightSeaGreen = FromArgb(255, 32, 178, 170);

        /// <summary>
        /// LightSkyBlue (Hex #FF87CEFA, ARGB: 255, 135, 206, 250)
        /// </summary>
        public static readonly Color LightSkyBlue = FromArgb(255, 135, 206, 250);

        /// <summary>
        /// LightSlateGray (Hex #FF778899, ARGB: 255, 119, 136, 153)
        /// </summary>
        public static readonly Color LightSlateGray = FromArgb(255, 119, 136, 153);

        /// <summary>
        /// LightSteelBlue (Hex #FFB0C4DE, ARGB: 255, 176, 196, 222)
        /// </summary>
        public static readonly Color LightSteelBlue = FromArgb(255, 176, 196, 222);

        /// <summary>
        /// LightYellow (Hex #FFFFFFE0, ARGB: 255, 255, 255, 224)
        /// </summary>
        public static readonly Color LightYellow = FromArgb(255, 255, 255, 224);

        /// <summary>
        /// Lime (Hex #FF00FF00, ARGB: 255, 0, 255, 0)
        /// </summary>
        public static readonly Color Lime = FromArgb(255, 0, 255, 0);

        /// <summary>
        /// LimeGreen (Hex #FF32CD32, ARGB: 255, 50, 205, 50)
        /// </summary>
        public static readonly Color LimeGreen = FromArgb(255, 50, 205, 50);

        /// <summary>
        /// Linen (Hex #FFFAF0E6, ARGB: 255, 250, 240, 230)
        /// </summary>
        public static readonly Color Linen = FromArgb(255, 250, 240, 230);

        /// <summary>
        /// Magenta (Hex #FFFF00FF, ARGB: 255, 255, 0, 255)
        /// </summary>
        public static readonly Color Magenta = FromArgb(255, 255, 0, 255);

        /// <summary>
        /// Maroon (Hex #FF800000, ARGB: 255, 128, 0, 0)
        /// </summary>
        public static readonly Color Maroon = FromArgb(255, 128, 0, 0);

        /// <summary>
        /// MediumAquamarine (Hex #FF66CDAA, ARGB: 255, 102, 205, 170)
        /// </summary>
        public static readonly Color MediumAquamarine = FromArgb(255, 102, 205, 170);

        /// <summary>
        /// MediumBlue (Hex #FF0000CD, ARGB: 255, 0, 0, 205)
        /// </summary>
        public static readonly Color MediumBlue = FromArgb(255, 0, 0, 205);

        /// <summary>
        /// MediumOrchid (Hex #FFBA55D3, ARGB: 255, 186, 85, 211)
        /// </summary>
        public static readonly Color MediumOrchid = FromArgb(255, 186, 85, 211);

        /// <summary>
        /// MediumPurple (Hex #FF9370DB, ARGB: 255, 147, 112, 219)
        /// </summary>
        public static readonly Color MediumPurple = FromArgb(255, 147, 112, 219);

        /// <summary>
        /// MediumSeaGreen (Hex #FF3CB371, ARGB: 255, 60, 179, 113)
        /// </summary>
        public static readonly Color MediumSeaGreen = FromArgb(255, 60, 179, 113);

        /// <summary>
        /// MediumSlateBlue (Hex #FF7B68EE, ARGB: 255, 123, 104, 238)
        /// </summary>
        public static readonly Color MediumSlateBlue = FromArgb(255, 123, 104, 238);

        /// <summary>
        /// MediumSpringGreen (Hex #FF00FA9A, ARGB: 255, 0, 250, 154)
        /// </summary>
        public static readonly Color MediumSpringGreen = FromArgb(255, 0, 250, 154);

        /// <summary>
        /// MediumTurquoise (Hex #FF48D1CC, ARGB: 255, 72, 209, 204)
        /// </summary>
        public static readonly Color MediumTurquoise = FromArgb(255, 72, 209, 204);

        /// <summary>
        /// MediumVioletRed (Hex #FFC71585, ARGB: 255, 199, 21, 133)
        /// </summary>
        public static readonly Color MediumVioletRed = FromArgb(255, 199, 21, 133);

        /// <summary>
        /// MidnightBlue (Hex #FF191970, ARGB: 255, 25, 25, 112)
        /// </summary>
        public static readonly Color MidnightBlue = FromArgb(255, 25, 25, 112);

        /// <summary>
        /// MintCream (Hex #FFF5FFFA, ARGB: 255, 245, 255, 250)
        /// </summary>
        public static readonly Color MintCream = FromArgb(255, 245, 255, 250);

        /// <summary>
        /// MistyRose (Hex #FFFFE4E1, ARGB: 255, 255, 228, 225)
        /// </summary>
        public static readonly Color MistyRose = FromArgb(255, 255, 228, 225);

        /// <summary>
        /// Moccasin (Hex #FFFFE4B5, ARGB: 255, 255, 228, 181)
        /// </summary>
        public static readonly Color Moccasin = FromArgb(255, 255, 228, 181);

        /// <summary>
        /// NavajoWhite (Hex #FFFFDEAD, ARGB: 255, 255, 222, 173)
        /// </summary>
        public static readonly Color NavajoWhite = FromArgb(255, 255, 222, 173);

        /// <summary>
        /// Navy (Hex #FF000080, ARGB: 255, 0, 0, 128)
        /// </summary>
        public static readonly Color Navy = FromArgb(255, 0, 0, 128);

        /// <summary>
        /// OldLace (Hex #FFFDF5E6, ARGB: 255, 253, 245, 230)
        /// </summary>
        public static readonly Color OldLace = FromArgb(255, 253, 245, 230);

        /// <summary>
        /// Olive (Hex #FF808000, ARGB: 255, 128, 128, 0)
        /// </summary>
        public static readonly Color Olive = FromArgb(255, 128, 128, 0);

        /// <summary>
        /// OliveDrab (Hex #FF6B8E23, ARGB: 255, 107, 142, 35)
        /// </summary>
        public static readonly Color OliveDrab = FromArgb(255, 107, 142, 35);

        /// <summary>
        /// Orange (Hex #FFFFA500, ARGB: 255, 255, 165, 0)
        /// </summary>
        public static readonly Color Orange = FromArgb(255, 255, 165, 0);

        /// <summary>
        /// OrangeRed (Hex #FFFF4500, ARGB: 255, 255, 69, 0)
        /// </summary>
        public static readonly Color OrangeRed = FromArgb(255, 255, 69, 0);

        /// <summary>
        /// Orchid (Hex #FFDA70D6, ARGB: 255, 218, 112, 214)
        /// </summary>
        public static readonly Color Orchid = FromArgb(255, 218, 112, 214);

        /// <summary>
        /// PaleGoldenrod (Hex #FFEEE8AA, ARGB: 255, 238, 232, 170)
        /// </summary>
        public static readonly Color PaleGoldenrod = FromArgb(255, 238, 232, 170);

        /// <summary>
        /// PaleGreen (Hex #FF98FB98, ARGB: 255, 152, 251, 152)
        /// </summary>
        public static readonly Color PaleGreen = FromArgb(255, 152, 251, 152);

        /// <summary>
        /// PaleTurquoise (Hex #FFAFEEEE, ARGB: 255, 175, 238, 238)
        /// </summary>
        public static readonly Color PaleTurquoise = FromArgb(255, 175, 238, 238);

        /// <summary>
        /// PaleVioletRed (Hex #FFDB7093, ARGB: 255, 219, 112, 147)
        /// </summary>
        public static readonly Color PaleVioletRed = FromArgb(255, 219, 112, 147);

        /// <summary>
        /// PapayaWhip (Hex #FFFFEFD5, ARGB: 255, 255, 239, 213)
        /// </summary>
        public static readonly Color PapayaWhip = FromArgb(255, 255, 239, 213);

        /// <summary>
        /// PeachPuff (Hex #FFFFDAB9, ARGB: 255, 255, 218, 185)
        /// </summary>
        public static readonly Color PeachPuff = FromArgb(255, 255, 218, 185);

        /// <summary>
        /// Peru (Hex #FFCD853F, ARGB: 255, 205, 133, 63)
        /// </summary>
        public static readonly Color Peru = FromArgb(255, 205, 133, 63);

        /// <summary>
        /// Pink (Hex #FFFFC0CB, ARGB: 255, 255, 192, 203)
        /// </summary>
        public static readonly Color Pink = FromArgb(255, 255, 192, 203);

        /// <summary>
        /// Plum (Hex #FFDDA0DD, ARGB: 255, 221, 160, 221)
        /// </summary>
        public static readonly Color Plum = FromArgb(255, 221, 160, 221);

        /// <summary>
        /// PowderBlue (Hex #FFB0E0E6, ARGB: 255, 176, 224, 230)
        /// </summary>
        public static readonly Color PowderBlue = FromArgb(255, 176, 224, 230);

        /// <summary>
        /// Purple (Hex #FF800080, ARGB: 255, 128, 0, 128)
        /// </summary>
        public static readonly Color Purple = FromArgb(255, 128, 0, 128);

        /// <summary>
        /// Red (Hex #FFFF0000, ARGB: 255, 255, 0, 0)
        /// </summary>
        public static readonly Color Red = FromArgb(255, 255, 0, 0);

        /// <summary>
        /// RosyBrown (Hex #FFBC8F8F, ARGB: 255, 188, 143, 143)
        /// </summary>
        public static readonly Color RosyBrown = FromArgb(255, 188, 143, 143);

        /// <summary>
        /// RoyalBlue (Hex #FF4169E1, ARGB: 255, 65, 105, 225)
        /// </summary>
        public static readonly Color RoyalBlue = FromArgb(255, 65, 105, 225);

        /// <summary>
        /// SaddleBrown (Hex #FF8B4513, ARGB: 255, 139, 69, 19)
        /// </summary>
        public static readonly Color SaddleBrown = FromArgb(255, 139, 69, 19);

        /// <summary>
        /// Salmon (Hex #FFFA8072, ARGB: 255, 250, 128, 114)
        /// </summary>
        public static readonly Color Salmon = FromArgb(255, 250, 128, 114);

        /// <summary>
        /// SandyBrown (Hex #FFF4A460, ARGB: 255, 244, 164, 96)
        /// </summary>
        public static readonly Color SandyBrown = FromArgb(255, 244, 164, 96);

        /// <summary>
        /// SeaGreen (Hex #FF2E8B57, ARGB: 255, 46, 139, 87)
        /// </summary>
        public static readonly Color SeaGreen = FromArgb(255, 46, 139, 87);

        /// <summary>
        /// SeaShell (Hex #FFFFF5EE, ARGB: 255, 255, 245, 238)
        /// </summary>
        public static readonly Color SeaShell = FromArgb(255, 255, 245, 238);

        /// <summary>
        /// Sienna (Hex #FFA0522D, ARGB: 255, 160, 82, 45)
        /// </summary>
        public static readonly Color Sienna = FromArgb(255, 160, 82, 45);

        /// <summary>
        /// Silver (Hex #FFC0C0C0, ARGB: 255, 192, 192, 192)
        /// </summary>
        public static readonly Color Silver = FromArgb(255, 192, 192, 192);

        /// <summary>
        /// SkyBlue (Hex #FF87CEEB, ARGB: 255, 135, 206, 235)
        /// </summary>
        public static readonly Color SkyBlue = FromArgb(255, 135, 206, 235);

        /// <summary>
        /// SlateBlue (Hex #FF6A5ACD, ARGB: 255, 106, 90, 205)
        /// </summary>
        public static readonly Color SlateBlue = FromArgb(255, 106, 90, 205);

        /// <summary>
        /// SlateGray (Hex #FF708090, ARGB: 255, 112, 128, 144)
        /// </summary>
        public static readonly Color SlateGray = FromArgb(255, 112, 128, 144);

        /// <summary>
        /// Snow (Hex #FFFFFAFA, ARGB: 255, 255, 250, 250)
        /// </summary>
        public static readonly Color Snow = FromArgb(255, 255, 250, 250);

        /// <summary>
        /// SpringGreen (Hex #FF00FF7F, ARGB: 255, 0, 255, 127)
        /// </summary>
        public static readonly Color SpringGreen = FromArgb(255, 0, 255, 127);

        /// <summary>
        /// SteelBlue (Hex #FF4682B4, ARGB: 255, 70, 130, 180)
        /// </summary>
        public static readonly Color SteelBlue = FromArgb(255, 70, 130, 180);

        /// <summary>
        /// Tan (Hex #FFD2B48C, ARGB: 255, 210, 180, 140)
        /// </summary>
        public static readonly Color Tan = FromArgb(255, 210, 180, 140);

        /// <summary>
        /// Teal (Hex #FF008080, ARGB: 255, 0, 128, 128)
        /// </summary>
        public static readonly Color Teal = FromArgb(255, 0, 128, 128);

        /// <summary>
        /// Thistle (Hex #FFD8BFD8, ARGB: 255, 216, 191, 216)
        /// </summary>
        public static readonly Color Thistle = FromArgb(255, 216, 191, 216);

        /// <summary>
        /// Tomato (Hex #FFFF6347, ARGB: 255, 255, 99, 71)
        /// </summary>
        public static readonly Color Tomato = FromArgb(255, 255, 99, 71);

        /// <summary>
        /// Transparent (Hex #00FFFFFF, ARGB: 0, 255, 255, 255)
        /// </summary>
        public static readonly Color Transparent = FromArgb(0, 255, 255, 255);

        /// <summary>
        /// Turquoise (Hex #FF40E0D0, ARGB: 255, 64, 224, 208)
        /// </summary>
        public static readonly Color Turquoise = FromArgb(255, 64, 224, 208);

        /// <summary>
        /// Violet (Hex #FFEE82EE, ARGB: 255, 238, 130, 238)
        /// </summary>
        public static readonly Color Violet = FromArgb(255, 238, 130, 238);

        /// <summary>
        /// Wheat (Hex #FFF5DEB3, ARGB: 255, 245, 222, 179)
        /// </summary>
        public static readonly Color Wheat = FromArgb(255, 245, 222, 179);

        /// <summary>
        /// White (Hex #FFFFFFFF, ARGB: 255, 255, 255, 255)
        /// </summary>
        public static readonly Color White = FromArgb(255, 255, 255, 255);

        /// <summary>
        /// WhiteSmoke (Hex #FFF5F5F5, ARGB: 255, 245, 245, 245)
        /// </summary>
        public static readonly Color WhiteSmoke = FromArgb(255, 245, 245, 245);

        /// <summary>
        /// Yellow (Hex #FFFFFF00, ARGB: 255, 255, 255, 0)
        /// </summary>
        public static readonly Color Yellow = FromArgb(255, 255, 255, 0);

        /// <summary>
        /// YellowGreen (Hex #FF9ACD32, ARGB: 255, 154, 205, 50)
        /// </summary>
        public static readonly Color YellowGreen = FromArgb(255, 154, 205, 50);
    }
}