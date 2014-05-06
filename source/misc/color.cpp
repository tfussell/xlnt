#include <algorithm>
#include <iomanip>
#include <sstream>
#include <vector>

#include "color.h"
#include "enum_iterator.h"

namespace xlnt {

const color color::AliceBlue = color::from_argb(255, 240, 248, 255);
const color color::AntiqueWhite = color::from_argb(255, 250, 235, 215);
const color color::Aqua = color::from_argb(255, 0, 255, 255);
const color color::Aquamarine = color::from_argb(255, 127, 255, 212);
const color color::Azure = color::from_argb(255, 240, 255, 255);
const color color::Beige = color::from_argb(255, 245, 245, 220);
const color color::Bisque = color::from_argb(255, 255, 228, 196);
const color color::Black = color::from_argb(255, 0, 0, 0);
const color color::BlanchedAlmond = color::from_argb(255, 255, 235, 205);
const color color::Blue = color::from_argb(255, 0, 0, 255);
const color color::BlueViolet = color::from_argb(255, 138, 43, 226);
const color color::Brown = color::from_argb(255, 165, 42, 42);
const color color::BurlyWood = color::from_argb(255, 222, 184, 135);
const color color::CadetBlue = color::from_argb(255, 95, 158, 160);
const color color::Chartreuse = color::from_argb(255, 127, 255, 0);
const color color::Chocolate = color::from_argb(255, 210, 105, 30);
const color color::Coral = color::from_argb(255, 255, 127, 80);
const color color::CornflowerBlue = color::from_argb(255, 100, 149, 237);
const color color::Cornsilk = color::from_argb(255, 255, 248, 220);
const color color::Crimson = color::from_argb(255, 220, 20, 60);
const color color::Cyan = color::from_argb(255, 0, 255, 255);
const color color::DarkBlue = color::from_argb(255, 0, 0, 139);
const color color::DarkCyan = color::from_argb(255, 0, 139, 139);
const color color::DarkGoldenrod = color::from_argb(255, 184, 134, 11);
const color color::DarkGray = color::from_argb(255, 169, 169, 169);
const color color::DarkGreen = color::from_argb(255, 0, 100, 0);
const color color::DarkKhaki = color::from_argb(255, 189, 183, 107);
const color color::DarkMagenta = color::from_argb(255, 139, 0, 139);
const color color::DarkOliveGreen = color::from_argb(255, 85, 107, 47);
const color color::DarkOrange = color::from_argb(255, 255, 140, 0);
const color color::DarkOrchid = color::from_argb(255, 153, 50, 204);
const color color::DarkRed = color::from_argb(255, 139, 0, 0);
const color color::DarkSalmon = color::from_argb(255, 233, 150, 122);
const color color::DarkSeaGreen = color::from_argb(255, 143, 188, 143);
const color color::DarkSlateBlue = color::from_argb(255, 72, 61, 139);
const color color::DarkSlateGray = color::from_argb(255, 47, 79, 79);
const color color::DarkTurquoise = color::from_argb(255, 0, 206, 209);
const color color::DarkViolet = color::from_argb(255, 148, 0, 211);
const color color::DeepPink = color::from_argb(255, 255, 20, 147);
const color color::DeepSkyBlue = color::from_argb(255, 0, 191, 255);
const color color::DimGray = color::from_argb(255, 105, 105, 105);
const color color::DodgerBlue = color::from_argb(255, 30, 144, 255);
const color color::Firebrick = color::from_argb(255, 178, 34, 34);
const color color::FloralWhite = color::from_argb(255, 255, 250, 240);
const color color::ForestGreen = color::from_argb(255, 34, 139, 34);
const color color::Fuchsia = color::from_argb(255, 255, 0, 255);
const color color::Gainsboro = color::from_argb(255, 220, 220, 220);
const color color::GhostWhite = color::from_argb(255, 248, 248, 255);
const color color::Gold = color::from_argb(255, 255, 215, 0);
const color color::Goldenrod = color::from_argb(255, 218, 165, 32);
const color color::Gray = color::from_argb(255, 128, 128, 128);
const color color::Green = color::from_argb(255, 0, 128, 0);
const color color::GreenYellow = color::from_argb(255, 173, 255, 47);
const color color::Honeydew = color::from_argb(255, 240, 255, 240);
const color color::HotPink = color::from_argb(255, 255, 105, 180);
const color color::IndianRed = color::from_argb(255, 205, 92, 92);
const color color::Indigo = color::from_argb(255, 75, 0, 130);
const color color::Ivory = color::from_argb(255, 255, 255, 240);
const color color::Khaki = color::from_argb(255, 240, 230, 140);
const color color::Lavender = color::from_argb(255, 230, 230, 250);
const color color::LavenderBlush = color::from_argb(255, 255, 240, 245);
const color color::LawnGreen = color::from_argb(255, 124, 252, 0);
const color color::LemonChiffon = color::from_argb(255, 255, 250, 205);
const color color::LightBlue = color::from_argb(255, 173, 216, 230);
const color color::LightCoral = color::from_argb(255, 240, 128, 128);
const color color::LightCyan = color::from_argb(255, 224, 255, 255);
const color color::LightGoldenrodYellow = color::from_argb(255, 250, 250, 210);
const color color::LightGray = color::from_argb(255, 211, 211, 211);
const color color::LightGreen = color::from_argb(255, 144, 238, 144);
const color color::LightPink = color::from_argb(255, 255, 182, 193);
const color color::LightSalmon = color::from_argb(255, 255, 160, 122);
const color color::LightSeaGreen = color::from_argb(255, 32, 178, 170);
const color color::LightSkyBlue = color::from_argb(255, 135, 206, 250);
const color color::LightSlateGray = color::from_argb(255, 119, 136, 153);
const color color::LightSteelBlue = color::from_argb(255, 176, 196, 222);
const color color::LightYellow = color::from_argb(255, 255, 255, 224);
const color color::Lime = color::from_argb(255, 0, 255, 0);
const color color::LimeGreen = color::from_argb(255, 50, 205, 50);
const color color::Linen = color::from_argb(255, 250, 240, 230);
const color color::Magenta = color::from_argb(255, 255, 0, 255);
const color color::Maroon = color::from_argb(255, 128, 0, 0);
const color color::MediumAquamarine = color::from_argb(255, 102, 205, 170);
const color color::MediumBlue = color::from_argb(255, 0, 0, 205);
const color color::MediumOrchid = color::from_argb(255, 186, 85, 211);
const color color::MediumPurple = color::from_argb(255, 147, 112, 219);
const color color::MediumSeaGreen = color::from_argb(255, 60, 179, 113);
const color color::MediumSlateBlue = color::from_argb(255, 123, 104, 238);
const color color::MediumSpringGreen = color::from_argb(255, 0, 250, 154);
const color color::MediumTurquoise = color::from_argb(255, 72, 209, 204);
const color color::MediumVioletRed = color::from_argb(255, 199, 21, 133);
const color color::MidnightBlue = color::from_argb(255, 25, 25, 112);
const color color::MintCream = color::from_argb(255, 245, 255, 250);
const color color::MistyRose = color::from_argb(255, 255, 228, 225);
const color color::Moccasin = color::from_argb(255, 255, 228, 181);
const color color::NavajoWhite = color::from_argb(255, 255, 222, 173);
const color color::Navy = color::from_argb(255, 0, 0, 128);
const color color::OldLace = color::from_argb(255, 253, 245, 230);
const color color::Olive = color::from_argb(255, 128, 128, 0);
const color color::OliveDrab = color::from_argb(255, 107, 142, 35);
const color color::Orange = color::from_argb(255, 255, 165, 0);
const color color::OrangeRed = color::from_argb(255, 255, 69, 0);
const color color::Orchid = color::from_argb(255, 218, 112, 214);
const color color::PaleGoldenrod = color::from_argb(255, 238, 232, 170);
const color color::PaleGreen = color::from_argb(255, 152, 251, 152);
const color color::PaleTurquoise = color::from_argb(255, 175, 238, 238);
const color color::PaleVioletRed = color::from_argb(255, 219, 112, 147);
const color color::PapayaWhip = color::from_argb(255, 255, 239, 213);
const color color::PeachPuff = color::from_argb(255, 255, 218, 185);
const color color::Peru = color::from_argb(255, 205, 133, 63);
const color color::Pink = color::from_argb(255, 255, 192, 203);
const color color::Plum = color::from_argb(255, 221, 160, 221);
const color color::PowderBlue = color::from_argb(255, 176, 224, 230);
const color color::Purple = color::from_argb(255, 128, 0, 128);
const color color::Red = color::from_argb(255, 255, 0, 0);
const color color::RosyBrown = color::from_argb(255, 188, 143, 143);
const color color::RoyalBlue = color::from_argb(255, 65, 105, 225);
const color color::SaddleBrown = color::from_argb(255, 139, 69, 19);
const color color::Salmon = color::from_argb(255, 250, 128, 114);
const color color::SandyBrown = color::from_argb(255, 244, 164, 96);
const color color::SeaGreen = color::from_argb(255, 46, 139, 87);
const color color::SeaShell = color::from_argb(255, 255, 245, 238);
const color color::Sienna = color::from_argb(255, 160, 82, 45);
const color color::Silver = color::from_argb(255, 192, 192, 192);
const color color::SkyBlue = color::from_argb(255, 135, 206, 235);
const color color::SlateBlue = color::from_argb(255, 106, 90, 205);
const color color::SlateGray = color::from_argb(255, 112, 128, 144);
const color color::Snow = color::from_argb(255, 255, 250, 250);
const color color::SpringGreen = color::from_argb(255, 0, 255, 127);
const color color::SteelBlue = color::from_argb(255, 70, 130, 180);
const color color::Tan = color::from_argb(255, 210, 180, 140);
const color color::Teal = color::from_argb(255, 0, 128, 128);
const color color::Thistle = color::from_argb(255, 216, 191, 216);
const color color::Tomato = color::from_argb(255, 255, 99, 71);
const color color::Transparent = color::from_argb(0, 0, 0, 0);
const color color::Turquoise = color::from_argb(255, 64, 224, 208);
const color color::Violet = color::from_argb(255, 238, 130, 238);
const color color::Wheat = color::from_argb(255, 245, 222, 179);
const color color::White = color::from_argb(255, 255, 255, 255);
const color color::WhiteSmoke = color::from_argb(255, 245, 245, 245);
const color color::Yellow = color::from_argb(255, 255, 255, 0);
const color color::YellowGreen = color::from_argb(255, 154, 205, 50);
const color color::Empty;

color color::from_argb(uint32_t argb)
{
	return color(uint8_t(argb >> 24), uint8_t(argb >> 16), uint8_t(argb >> 8), uint8_t(argb));
}

color color::from_argb(uint8_t alpha, color base)
{
	return color(alpha, base.red_, base.green_, base.blue_);
}

color color::from_argb(uint8_t r, uint8_t g, uint8_t b)
{
	return color(255, r, g, b);
}

color color::from_argb(uint8_t a, uint8_t r, uint8_t g, uint8_t b)
{
	return color(a, r, g, b);
}

color color::from_known_color(known_color color)
{
	switch(color)
	{
	case known_color::AliceBlue: return AliceBlue;
	case known_color::AntiqueWhite: return AntiqueWhite;
	case known_color::Aqua: return Aqua;
	case known_color::Aquamarine: return Aquamarine;
	case known_color::Azure: return Azure;
	case known_color::Beige: return Beige;
	case known_color::Bisque: return Bisque;
	case known_color::Black: return Black;
	case known_color::BlanchedAlmond: return BlanchedAlmond;
	case known_color::Blue: return Blue;
	case known_color::BlueViolet: return BlueViolet;
	case known_color::Brown: return Brown;
	case known_color::BurlyWood: return BurlyWood;
	case known_color::CadetBlue: return CadetBlue;
	case known_color::Chartreuse: return Chartreuse;
	case known_color::Chocolate: return Chocolate;
	case known_color::Coral: return Coral;
	case known_color::CornflowerBlue: return CornflowerBlue;
	case known_color::Cornsilk: return Cornsilk;
	case known_color::Crimson: return Crimson;
	case known_color::Cyan: return Cyan;
	case known_color::DarkBlue: return DarkBlue;
	case known_color::DarkCyan: return DarkCyan;
	case known_color::DarkGoldenrod: return DarkGoldenrod;
	case known_color::DarkGray: return DarkGray;
	case known_color::DarkGreen: return DarkGreen;
	case known_color::DarkKhaki: return DarkKhaki;
	case known_color::DarkMagenta: return DarkMagenta;
	case known_color::DarkOliveGreen: return DarkOliveGreen;
	case known_color::DarkOrange: return DarkOrange;
	case known_color::DarkOrchid: return DarkOrchid;
	case known_color::DarkRed: return DarkRed;
	case known_color::DarkSalmon: return DarkSalmon;
	case known_color::DarkSeaGreen: return DarkSeaGreen;
	case known_color::DarkSlateBlue: return DarkSlateBlue;
	case known_color::DarkSlateGray: return DarkSlateGray;
	case known_color::DarkTurquoise: return DarkTurquoise;
	case known_color::DarkViolet: return DarkViolet;
	case known_color::DeepPink: return DeepPink;
	case known_color::DeepSkyBlue: return DeepSkyBlue;
	case known_color::DimGray: return DimGray;
	case known_color::DodgerBlue: return DodgerBlue;
	case known_color::Firebrick: return Firebrick;
	case known_color::FloralWhite: return FloralWhite;
	case known_color::ForestGreen: return ForestGreen;
	case known_color::Fuchsia: return Fuchsia;
	case known_color::Gainsboro: return Gainsboro;
	case known_color::GhostWhite: return GhostWhite;
	case known_color::Gold: return Gold;
	case known_color::Goldenrod: return Goldenrod;
	case known_color::Gray: return Gray;
	case known_color::Green: return Green;
	case known_color::GreenYellow: return GreenYellow;
	case known_color::Honeydew: return Honeydew;
	case known_color::HotPink: return HotPink;
	case known_color::IndianRed: return IndianRed;
	case known_color::Indigo: return Indigo;
	case known_color::Ivory: return Ivory;
	case known_color::Khaki: return Khaki;
	case known_color::Lavender: return Lavender;
	case known_color::LavenderBlush: return LavenderBlush;
	case known_color::LawnGreen: return LawnGreen;
	case known_color::LemonChiffon: return LemonChiffon;
	case known_color::LightBlue: return LightBlue;
	case known_color::LightCoral: return LightCoral;
	case known_color::LightCyan: return LightCyan;
	case known_color::LightGoldenrodYellow: return LightGoldenrodYellow;
	case known_color::LightGray: return LightGray;
	case known_color::LightGreen: return LightGreen;
	case known_color::LightPink: return LightPink;
	case known_color::LightSalmon: return LightSalmon;
	case known_color::LightSeaGreen: return LightSeaGreen;
	case known_color::LightSkyBlue: return LightSkyBlue;
	case known_color::LightSlateGray: return LightSlateGray;
	case known_color::LightSteelBlue: return LightSteelBlue;
	case known_color::LightYellow: return LightYellow;
	case known_color::Lime: return Lime;
	case known_color::LimeGreen: return LimeGreen;
	case known_color::Linen: return Linen;
	case known_color::Magenta: return Magenta;
	case known_color::Maroon: return Maroon;
	case known_color::MediumAquamarine: return MediumAquamarine;
	case known_color::MediumBlue: return MediumBlue;
	case known_color::MediumOrchid: return MediumOrchid;
	case known_color::MediumPurple: return MediumPurple;
	case known_color::MediumSeaGreen: return MediumSeaGreen;
	case known_color::MediumSlateBlue: return MediumSlateBlue;
	case known_color::MediumSpringGreen: return MediumSpringGreen;
	case known_color::MediumTurquoise: return MediumTurquoise;
	case known_color::MediumVioletRed: return MediumVioletRed;
	case known_color::MidnightBlue: return MidnightBlue;
	case known_color::MintCream: return MintCream;
	case known_color::MistyRose: return MistyRose;
	case known_color::Moccasin: return Moccasin;
	case known_color::NavajoWhite: return NavajoWhite;
	case known_color::Navy: return Navy;
	case known_color::OldLace: return OldLace;
	case known_color::Olive: return Olive;
	case known_color::OliveDrab: return OliveDrab;
	case known_color::Orange: return Orange;
	case known_color::OrangeRed: return OrangeRed;
	case known_color::Orchid: return Orchid;
	case known_color::PaleGoldenrod: return PaleGoldenrod;
	case known_color::PaleGreen: return PaleGreen;
	case known_color::PaleTurquoise: return PaleTurquoise;
	case known_color::PaleVioletRed: return PaleVioletRed;
	case known_color::PapayaWhip: return PapayaWhip;
	case known_color::PeachPuff: return PeachPuff;
	case known_color::Peru: return Peru;
	case known_color::Pink: return Pink;
	case known_color::Plum: return Plum;
	case known_color::PowderBlue: return PowderBlue;
	case known_color::Purple: return Purple;
	case known_color::Red: return Red;
	case known_color::RosyBrown: return RosyBrown;
	case known_color::RoyalBlue: return RoyalBlue;
	case known_color::SaddleBrown: return SaddleBrown;
	case known_color::Salmon: return Salmon;
	case known_color::SandyBrown: return SandyBrown;
	case known_color::SeaGreen: return SeaGreen;
	case known_color::SeaShell: return SeaShell;
	case known_color::Sienna: return Sienna;
	case known_color::Silver: return Silver;
	case known_color::SkyBlue: return SkyBlue;
	case known_color::SlateBlue: return SlateBlue;
	case known_color::SlateGray: return SlateGray;
	case known_color::Snow: return Snow;
	case known_color::SpringGreen: return SpringGreen;
	case known_color::SteelBlue: return SteelBlue;
	case known_color::Tan: return Tan;
	case known_color::Teal: return Teal;
	case known_color::Thistle: return Thistle;
	case known_color::Tomato: return Tomato;
	case known_color::Turquoise: return Turquoise;
	case known_color::Violet: return Violet;
	case known_color::Wheat: return Wheat;
	case known_color::White: return White;
	case known_color::WhiteSmoke: return WhiteSmoke;
	case known_color::Yellow: return Yellow;
	case known_color::YellowGreen: return YellowGreen;
	}
	throw std::runtime_error("unknown known_color");
}

color color::from_name(const std::string &name)
{
	for(auto current_color : enum_iterator<known_color>())
	{
		if(to_string(current_color) == name)
		{
			return from_known_color(current_color);
		}
	}
	return Empty;
}

bool color::equals(color b) const
{
	if(is_empty_)
	{
		return b.is_empty_;
	}

	return !b.is_empty_ && alpha_ == b.alpha_ && red_ == b.red_ && green_ == b.green_ && blue_ == b.blue_;
}

struct hsv
{
	float h;
	float s;
	float v;
};

hsv RGBtoHSV(float red, float green, float blue)
{
	hsv r;
	float min = std::min({red, green, blue});
	float max = std::max({red, green, blue});
	r.v = max;				// v
	float delta = max - min;
	if(max != 0)
	{
		r.s = delta / max;		// s
	}
	else 
	{
		// r = g = b = 0		// s = 0, v is undefined
		r.s = 0;
		r.h = -1;
		return r;
	}
	if(red == max)
	{
		r.h = (green - blue) / delta;		// between yellow & magenta
	}
	else if(green == max)
	{
		r.h = 2 + (blue - red) / delta;	// between cyan & yellow
	}
	else
	{
		r.h = 4 + (red - green) / delta;	// between magenta & cyan
	}
	r.h *= 60;				// degrees
	if(r.h < 0)
	{
		r.h += 360;
	}
	return r;
}

float color::get_brightness() const
{
	return RGBtoHSV(red_, green_, blue_).v;
}

float color::get_hue() const
{
	return RGBtoHSV(red_, green_, blue_).h;
}

float color::getsaturation() const
{
	return RGBtoHSV(red_, green_, blue_).s;
}

bool color::is_known_color() const
{
	try
	{
		to_known_color();
		return true;
	}
	catch(...)
	{
		return false;
	}
}

bool color::is_system_color() const
{
	static const std::vector<known_color> system_colors = {
		known_color::ActiveBorder,
		known_color::ActiveCaption,
		known_color::ActiveCaptionText,
		known_color::AppWorkspace,
		known_color::ButtonFace,
		known_color::ButtonHighlight,
		known_color::ButtonShadow,
		known_color::Control,
		known_color::ControlDark,
		known_color::ControlDarkDark,
		known_color::ControlLight,
		known_color::ControlLightLight,
		known_color::ControlText,
		known_color::Desktop,
		known_color::GradientActiveCaption,
		known_color::GradientInactiveCaption,
		known_color::GrayText,
		known_color::Highlight,
		known_color::HighlightText,
		known_color::HotTrack,
		known_color::InactiveBorder,
		known_color::InactiveCaption,
		known_color::InactiveCaptionText,
		known_color::Info,
		known_color::InfoText,
		known_color::Menu,
		known_color::MenuBar,
		known_color::MenuHighlight,
		known_color::MenuText,
		known_color::ScrollBar,
		known_color::Window,
		known_color::WindowFrame,
		known_color::WindowText
	};

	try
	{
		return std::find(system_colors.begin(), system_colors.end(), to_known_color()) != system_colors.end();
	}
	catch(...)
	{
		return false;
	}
}

uint32_t color::to_argb() const
{
	uint32_t argb = alpha_;
	argb = argb << 8 | red_;
	argb = argb << 8 | green_;
	argb = argb << 8 | blue_;
	return argb;
}

known_color color::to_known_color() const
{
	for(auto current_color : enum_iterator<known_color>())
	{
        if(from_known_color(current_color) == *this)
		{
            return current_color;
		}
	}

	return known_color::Unknown;
}

std::string color::get_name() const
{
	try
	{
		return to_string(to_known_color());
	}
	catch(...)
	{
		std::stringstream stream;
		stream << "#" << std::setfill('0') << std::setw(2) << std::hex;
		stream << alpha_ << red_ << green_ << blue_;
		return stream.str();
	}
}

} // namespace xlnt
