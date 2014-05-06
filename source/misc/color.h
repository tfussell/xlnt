#pragma once

#include <cstdint>
#include "enum_iterator.h"

namespace xlnt {

enum class known_color
{
	/// <summary>
	/// This is returned when a color isn't known.
	/// </summary>
	Unknown = 0,
	/// <summary>
	/// The system-defined color of the active window's border.
	/// </summary>
	ActiveBorder, 
	/// <summary>
	/// The system-defined color of the background of the active window's title bar.
	/// </summary>
	ActiveCaption, 
	/// <summary>
	/// The system-defined color of the text in the active window's title bar.
	/// </summary>
	ActiveCaptionText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	AliceBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	AntiqueWhite, 
	/// <summary>
	/// The system-defined color of the application workspace. The application workspace is the area in a multiple-document view that is not being occupied by documents.
	/// </summary>
	AppWorkspace, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Aqua, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Aquamarine, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Azure, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Beige, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Bisque, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Black, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	BlanchedAlmond, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Blue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	BlueViolet, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Brown, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	BurlyWood, 
	/// <summary>
	/// The system-defined face color of a 3-D element.
	/// </summary>
	ButtonFace, 
	/// <summary>
	/// The system-defined color that is the highlight color of a 3-D element. This color is applied to parts of a 3-D element that face the light source.
	/// </summary>
	ButtonHighlight, 
	/// <summary>
	/// The system-defined color that is the shadow color of a 3-D element. This color is applied to parts of a 3-D element that face away from the light source.
	/// </summary>
	ButtonShadow, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	CadetBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Chartreuse, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Chocolate, 
	/// <summary>
	/// The system-defined face color of a 3-D element.
	/// </summary>
	Control, 
	/// <summary>
	/// The system-defined shadow color of a 3-D element. The shadow color is applied to parts of a 3-D element that face away from the light source.
	/// </summary>
	ControlDark, 
	/// <summary>
	/// The system-defined color that is the dark shadow color of a 3-D element. The dark shadow color is applied to the parts of a 3-D element that are the darkest color.
	/// </summary>
	ControlDarkDark, 
	/// <summary>
	/// The system-defined color that is the light color of a 3-D element. The light color is applied to parts of a 3-D element that face the light source.
	/// </summary>
	ControlLight, 
	/// <summary>
	/// The system-defined highlight color of a 3-D element. The highlight color is applied to the parts of a 3-D element that are the lightest color.
	/// </summary>
	ControlLightLight, 
	/// <summary>
	/// The system-defined color of text in a 3-D element.
	/// </summary>
	ControlText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Coral, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	CornflowerBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Cornsilk, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Crimson, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Cyan, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkCyan, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkGoldenrod, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkKhaki, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkMagenta, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkOliveGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkOrange, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkOrchid, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkRed, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkSalmon, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkSeaGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkSlateBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkSlateGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkTurquoise, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DarkViolet, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DeepPink, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DeepSkyBlue, 
	/// <summary>
	/// The system-defined color of the desktop.
	/// </summary>
	Desktop, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DimGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	DodgerBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Firebrick, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	FloralWhite, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	ForestGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Fuchsia, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Gainsboro, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	GhostWhite, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Gold, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Goldenrod, 
	/// <summary>
	/// The system-defined color of the lightest color in the color gradient of an active window's title bar.
	/// </summary>
	GradientActiveCaption, 
	/// <summary>
	/// The system-defined color of the lightest color in the color gradient of an inactive window's title bar.
	/// </summary>
	GradientInactiveCaption, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Gray, 
	/// <summary>
	/// The system-defined color of dimmed text. Items in a list that are disabled are displayed in dimmed text.
	/// </summary>
	GrayText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Green, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	GreenYellow, 
	/// <summary>
	/// The system-defined color of the background of selected items. This includes selected menu items as well as selected text.
	/// </summary>
	Highlight, 
	/// <summary>
	/// The system-defined color of the text of selected items.
	/// </summary>
	HighlightText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Honeydew, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	HotPink, 
	/// <summary>
	/// The system-defined color used to designate a hot-tracked item. Single-clicking a hot-tracked item executes the item.
	/// </summary>
	HotTrack, 
	/// <summary>
	/// The system-defined color of an inactive window's border.
	/// </summary>
	InactiveBorder, 
	/// <summary>
	/// The system-defined color of the background of an inactive window's title bar.
	/// </summary>
	InactiveCaption, 
	/// <summary>
	/// The system-defined color of the text in an inactive window's title bar.
	/// </summary>
	InactiveCaptionText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	IndianRed, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Indigo, 
	/// <summary>
	/// The system-defined color of the background of a ToolTip.
	/// </summary>
	Info, 
	/// <summary>
	/// The system-defined color of the text of a ToolTip.
	/// </summary>
	InfoText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Ivory, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Khaki, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Lavender, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LavenderBlush, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LawnGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LemonChiffon, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightCoral, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightCyan, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightGoldenrodYellow, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightPink, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightSalmon, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightSeaGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightSkyBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightSlateGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightSteelBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LightYellow, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Lime, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	LimeGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Linen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Magenta, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Maroon, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumAquamarine, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumOrchid, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumPurple, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumSeaGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumSlateBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumSpringGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumTurquoise, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MediumVioletRed, 
	/// <summary>
	/// The system-defined color of a menu's background.
	/// </summary>
	Menu, 
	/// <summary>
	/// The system-defined color of the background of a menu bar.
	/// </summary>
	MenuBar, 
	/// <summary>
	/// The system-defined color used to highlight menu items when the menu appears as a flat menu.
	/// </summary>
	MenuHighlight, 
	/// <summary>
	/// The system-defined color of a menu's text.
	/// </summary>
	MenuText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MidnightBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MintCream, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	MistyRose, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Moccasin, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	NavajoWhite, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Navy, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	OldLace, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Olive, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	OliveDrab, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Orange, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	OrangeRed, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Orchid, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PaleGoldenrod, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PaleGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PaleTurquoise, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PaleVioletRed, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PapayaWhip, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PeachPuff, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Peru, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Pink, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Plum, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	PowderBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Purple, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Red, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	RosyBrown, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	RoyalBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SaddleBrown, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Salmon, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SandyBrown, 
	/// <summary>
	/// The system-defined color of the background of a scroll bar.
	/// </summary>
	ScrollBar, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SeaGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SeaShell, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Sienna, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Silver, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SkyBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SlateBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SlateGray, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Snow, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SpringGreen, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	SteelBlue, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Tan, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Teal, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Thistle, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Tomato, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Transparent, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Turquoise, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Violet, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Wheat, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	White, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	WhiteSmoke, 
	/// <summary>
	/// The system-defined color of the background in the client area of a window.
	/// </summary>
	Window, 
	/// <summary>
	/// The system-defined color of a window frame.
	/// </summary>
	WindowFrame, 
	/// <summary>
	/// The system-defined color of the text in the client area of a window.
	/// </summary>
	WindowText, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	Yellow, 
	/// <summary>
	/// A system-defined color.
	/// </summary>
	YellowGreen, 
	Last,
	First = ActiveBorder
};

inline std::string to_string(known_color color)
{
	switch(color)
	{
	case known_color::ActiveBorder: return "ActiveBorder";
	case known_color::ActiveCaption: return "ActiveCaption";
	case known_color::ActiveCaptionText: return "ActiveCaptionText";
	case known_color::AliceBlue: return "AliceBlue";
	case known_color::AntiqueWhite: return "AntiqueWhite";
	case known_color::AppWorkspace: return "AppWorkspace";
	case known_color::Aqua: return "Aqua";
	case known_color::Aquamarine: return "Aquamarine";
	case known_color::Azure: return "Azure";
	case known_color::Beige: return "Beige";
	case known_color::Bisque: return "Bisque";
	case known_color::Black: return "Black";
	case known_color::BlanchedAlmond: return "BlanchedAlmond";
	case known_color::Blue: return "Blue";
	case known_color::BlueViolet: return "BlueViolet";
	case known_color::Brown: return "Brown";
	case known_color::BurlyWood: return "BurlyWood";
	case known_color::ButtonFace: return "ButtonFace";
	case known_color::ButtonHighlight: return "ButtonHighlight";
	case known_color::ButtonShadow: return "ButtonShadow";
	case known_color::CadetBlue: return "CadetBlue";
	case known_color::Chartreuse: return "Chartreuse";
	case known_color::Chocolate: return "Chocolate";
	case known_color::Control: return "Control";
	case known_color::ControlDark: return "ControlDark";
	case known_color::ControlDarkDark: return "ControlDarkDark";
	case known_color::ControlLight: return "ControlLight";
	case known_color::ControlLightLight: return "ControlLightLight";
	case known_color::ControlText: return "ControlText";
	case known_color::Coral: return "Coral";
	case known_color::CornflowerBlue: return "CornflowerBlue";
	case known_color::Cornsilk: return "Cornsilk";
	case known_color::Crimson: return "Crimson";
	case known_color::Cyan: return "Cyan";
	case known_color::DarkBlue: return "DarkBlue";
	case known_color::DarkCyan: return "DarkCyan";
	case known_color::DarkGoldenrod: return "DarkGoldenrod";
	case known_color::DarkGray: return "DarkGray";
	case known_color::DarkGreen: return "DarkGreen";
	case known_color::DarkKhaki: return "DarkKhaki";
	case known_color::DarkMagenta: return "DarkMagenta";
	case known_color::DarkOliveGreen: return "DarkOliveGreen";
	case known_color::DarkOrange: return "DarkOrange";
	case known_color::DarkOrchid: return "DarkOrchid";
	case known_color::DarkRed: return "DarkRed";
	case known_color::DarkSalmon: return "DarkSalmon";
	case known_color::DarkSeaGreen: return "DarkSeaGreen";
	case known_color::DarkSlateBlue: return "DarkSlateBlue";
	case known_color::DarkSlateGray: return "DarkSlateGray";
	case known_color::DarkTurquoise: return "DarkTurquoise";
	case known_color::DarkViolet: return "DarkViolet";
	case known_color::DeepPink: return "DeepPink";
	case known_color::DeepSkyBlue: return "DeepSkyBlue";
	case known_color::Desktop: return "Desktop";
	case known_color::DimGray: return "DimGray";
	case known_color::DodgerBlue: return "DodgerBlue";
	case known_color::Firebrick: return "Firebrick";
	case known_color::FloralWhite: return "FloralWhite";
	case known_color::ForestGreen: return "ForestGreen";
	case known_color::Fuchsia: return "Fuchsia";
	case known_color::Gainsboro: return "Gainsboro";
	case known_color::GhostWhite: return "GhostWhite";
	case known_color::Gold: return "Gold";
	case known_color::Goldenrod: return "Goldenrod";
	case known_color::GradientActiveCaption: return "GradientActiveCaption";
	case known_color::GradientInactiveCaption: return "GradientInactiveCaption";
	case known_color::Gray: return "Gray";
	case known_color::GrayText: return "GrayText";
	case known_color::Green: return "Green";
	case known_color::GreenYellow: return "GreenYellow";
	case known_color::Highlight: return "Highlight";
	case known_color::HighlightText: return "HighlightText";
	case known_color::Honeydew: return "Honeydew";
	case known_color::HotPink: return "HotPink";
	case known_color::HotTrack: return "HotTrack";
	case known_color::InactiveBorder: return "InactiveBorder";
	case known_color::InactiveCaption: return "InactiveCaption";
	case known_color::InactiveCaptionText: return "InactiveCaptionText";
	case known_color::IndianRed: return "IndianRed";
	case known_color::Indigo: return "Indigo";
	case known_color::Info: return "Info";
	case known_color::InfoText: return "InfoText";
	case known_color::Ivory: return "Ivory";
	case known_color::Khaki: return "Khaki";
	case known_color::Lavender: return "Lavender";
	case known_color::LavenderBlush: return "LavenderBlush";
	case known_color::LawnGreen: return "LawnGreen";
	case known_color::LemonChiffon: return "LemonChiffon";
	case known_color::LightBlue: return "LightBlue";
	case known_color::LightCoral: return "LightCoral";
	case known_color::LightCyan: return "LightCyan";
	case known_color::LightGoldenrodYellow: return "LightGoldenrodYellow";
	case known_color::LightGray: return "LightGray";
	case known_color::LightGreen: return "LightGreen";
	case known_color::LightPink: return "LightPink";
	case known_color::LightSalmon: return "LightSalmon";
	case known_color::LightSeaGreen: return "LightSeaGreen";
	case known_color::LightSkyBlue: return "LightSkyBlue";
	case known_color::LightSlateGray: return "LightSlateGray";
	case known_color::LightSteelBlue: return "LightSteelBlue";
	case known_color::LightYellow: return "LightYellow";
	case known_color::Lime: return "Lime";
	case known_color::LimeGreen: return "LimeGreen";
	case known_color::Linen: return "Linen";
	case known_color::Magenta: return "Magenta";
	case known_color::Maroon: return "Maroon";
	case known_color::MediumAquamarine: return "MediumAquamarine";
	case known_color::MediumBlue: return "MediumBlue";
	case known_color::MediumOrchid: return "MediumOrchid";
	case known_color::MediumPurple: return "MediumPurple";
	case known_color::MediumSeaGreen: return "MediumSeaGreen";
	case known_color::MediumSlateBlue: return "MediumSlateBlue";
	case known_color::MediumSpringGreen: return "MediumSpringGreen";
	case known_color::MediumTurquoise: return "MediumTurquoise";
	case known_color::MediumVioletRed: return "MediumVioletRed";
	case known_color::Menu: return "Menu";
	case known_color::MenuBar: return "MenuBar";
	case known_color::MenuHighlight: return "MenuHighlight";
	case known_color::MenuText: return "MenuText";
	case known_color::MidnightBlue: return "MidnightBlue";
	case known_color::MintCream: return "MintCream";
	case known_color::MistyRose: return "MistyRose";
	case known_color::Moccasin: return "Moccasin";
	case known_color::NavajoWhite: return "NavajoWhite";
	case known_color::Navy: return "Navy";
	case known_color::OldLace: return "OldLace";
	case known_color::Olive: return "Olive";
	case known_color::OliveDrab: return "OliveDrab";
	case known_color::Orange: return "Orange";
	case known_color::OrangeRed: return "OrangeRed";
	case known_color::Orchid: return "Orchid";
	case known_color::PaleGoldenrod: return "PaleGoldenrod";
	case known_color::PaleGreen: return "PaleGreen";
	case known_color::PaleTurquoise: return "PaleTurquoise";
	case known_color::PaleVioletRed: return "PaleVioletRed";
	case known_color::PapayaWhip: return "PapayaWhip";
	case known_color::PeachPuff: return "PeachPuff";
	case known_color::Peru: return "Peru";
	case known_color::Pink: return "Pink";
	case known_color::Plum: return "Plum";
	case known_color::PowderBlue: return "PowderBlue";
	case known_color::Purple: return "Purple";
	case known_color::Red: return "Red";
	case known_color::RosyBrown: return "RosyBrown";
	case known_color::RoyalBlue: return "RoyalBlue";
	case known_color::SaddleBrown: return "SaddleBrown";
	case known_color::Salmon: return "Salmon";
	case known_color::SandyBrown: return "SandyBrown";
	case known_color::ScrollBar: return "ScrollBar";
	case known_color::SeaGreen: return "SeaGreen";
	case known_color::SeaShell: return "SeaShell";
	case known_color::Sienna: return "Sienna";
	case known_color::Silver: return "Silver";
	case known_color::SkyBlue: return "SkyBlue";
	case known_color::SlateBlue: return "SlateBlue";
	case known_color::SlateGray: return "SlateGray";
	case known_color::Snow: return "Snow";
	case known_color::SpringGreen: return "SpringGreen";
	case known_color::SteelBlue: return "SteelBlue";
	case known_color::Tan: return "Tan";
	case known_color::Teal: return "Teal";
	case known_color::Thistle: return "Thistle";
	case known_color::Tomato: return "Tomato";
	case known_color::Transparent: return "Transparent";
	case known_color::Turquoise: return "Turquoise";
	case known_color::Violet: return "Violet";
	case known_color::Wheat: return "Wheat";
	case known_color::White: return "White";
	case known_color::WhiteSmoke: return "WhiteSmoke";
	case known_color::Window: return "Window";
	case known_color::WindowFrame: return "WindowFrame";
	case known_color::WindowText: return "WindowText";
	case known_color::Yellow: return "Yellow";
	case known_color::YellowGreen: return "YellowGreen";
	}
	throw std::runtime_error("");
}

class color
{
public:
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF0F8FF.
	/// </summary>
	static const color AliceBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFAEBD7.
	/// </summary>
	static const color AntiqueWhite;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0FFFF.
	/// </summary>
	static const color Aqua;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF7FFFD4.
	/// </summary>
	static const color Aquamarine;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF0FFFF.
	/// </summary>
	static const color Azure;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF5F5DC.
	/// </summary>
	static const color Beige;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFE4C4.
	/// </summary>
	static const color Bisque;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF000.
	/// </summary>
	static const color Black;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFEBCD.
	/// </summary>
	static const color BlanchedAlmond;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF00FF.
	/// </summary>
	static const color Blue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8A2BE2.
	/// </summary>
	static const color BlueViolet;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFA52A2A.
	/// </summary>
	static const color Brown;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDEB887.
	/// </summary>
	static const color BurlyWood;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF5F9EA0.
	/// </summary>
	static const color CadetBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF7FFF0.
	/// </summary>
	static const color Chartreuse;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFD2691E.
	/// </summary>
	static const color Chocolate;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF7F50.
	/// </summary>
	static const color Coral;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF6495ED.
	/// </summary>
	static const color CornflowerBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFF8DC.
	/// </summary>
	static const color Cornsilk;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDC143C.
	/// </summary>
	static const color Crimson;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0FFFF.
	/// </summary>
	static const color Cyan;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF008B.
	/// </summary>
	static const color DarkBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF08B8B.
	/// </summary>
	static const color DarkCyan;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFB8860B.
	/// </summary>
	static const color DarkGoldenrod;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFA9A9A9.
	/// </summary>
	static const color DarkGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0640.
	/// </summary>
	static const color DarkGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFBDB76B.
	/// </summary>
	static const color DarkKhaki;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8B08B.
	/// </summary>
	static const color DarkMagenta;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF556B2F.
	/// </summary>
	static const color DarkOliveGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF8C0.
	/// </summary>
	static const color DarkOrange;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF9932CC.
	/// </summary>
	static const color DarkOrchid;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8B00.
	/// </summary>
	static const color DarkRed;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFE9967A.
	/// </summary>
	static const color DarkSalmon;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8FBC8F.
	/// </summary>
	static const color DarkSeaGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF483D8B.
	/// </summary>
	static const color DarkSlateBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF2F4F4F.
	/// </summary>
	static const color DarkSlateGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0CED1.
	/// </summary>
	static const color DarkTurquoise;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF940D3.
	/// </summary>
	static const color DarkViolet;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF1493.
	/// </summary>
	static const color DeepPink;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0BFFF.
	/// </summary>
	static const color DeepSkyBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF696969.
	/// </summary>
	static const color DimGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF1E90FF.
	/// </summary>
	static const color DodgerBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFB22222.
	/// </summary>
	static const color Firebrick;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFAF0.
	/// </summary>
	static const color FloralWhite;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF228B22.
	/// </summary>
	static const color ForestGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF0FF.
	/// </summary>
	static const color Fuchsia;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDCDCDC.
	/// </summary>
	static const color Gainsboro;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF8F8FF.
	/// </summary>
	static const color GhostWhite;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFD70.
	/// </summary>
	static const color Gold;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDAA520.
	/// </summary>
	static const color Goldenrod;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF808080.
	/// </summary>
	static const color Gray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0800.
	/// </summary>
	static const color Green;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFADFF2F.
	/// </summary>
	static const color GreenYellow;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF0FFF0.
	/// </summary>
	static const color Honeydew;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF69B4.
	/// </summary>
	static const color HotPink;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFCD5C5C.
	/// </summary>
	static const color IndianRed;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF4B082.
	/// </summary>
	static const color Indigo;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFFF0.
	/// </summary>
	static const color Ivory;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF0E68C.
	/// </summary>
	static const color Khaki;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFE6E6FA.
	/// </summary>
	static const color Lavender;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFF0F5.
	/// </summary>
	static const color LavenderBlush;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF7CFC0.
	/// </summary>
	static const color LawnGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFACD.
	/// </summary>
	static const color LemonChiffon;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFADD8E6.
	/// </summary>
	static const color LightBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF08080.
	/// </summary>
	static const color LightCoral;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFE0FFFF.
	/// </summary>
	static const color LightCyan;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFAFAD2.
	/// </summary>
	static const color LightGoldenrodYellow;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFD3D3D3.
	/// </summary>
	static const color LightGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF90EE90.
	/// </summary>
	static const color LightGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFB6C1.
	/// </summary>
	static const color LightPink;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFA07A.
	/// </summary>
	static const color LightSalmon;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF20B2AA.
	/// </summary>
	static const color LightSeaGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF87CEFA.
	/// </summary>
	static const color LightSkyBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF778899.
	/// </summary>
	static const color LightSlateGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFB0C4DE.
	/// </summary>
	static const color LightSteelBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFFE0.
	/// </summary>
	static const color LightYellow;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0FF0.
	/// </summary>
	static const color Lime;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF32CD32.
	/// </summary>
	static const color LimeGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFAF0E6.
	/// </summary>
	static const color Linen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF0FF.
	/// </summary>
	static const color Magenta;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8000.
	/// </summary>
	static const color Maroon;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF66CDAA.
	/// </summary>
	static const color MediumAquamarine;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF00CD.
	/// </summary>
	static const color MediumBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFBA55D3.
	/// </summary>
	static const color MediumOrchid;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF9370DB.
	/// </summary>
	static const color MediumPurple;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF3CB371.
	/// </summary>
	static const color MediumSeaGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF7B68EE.
	/// </summary>
	static const color MediumSlateBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0FA9A.
	/// </summary>
	static const color MediumSpringGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF48D1CC.
	/// </summary>
	static const color MediumTurquoise;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFC71585.
	/// </summary>
	static const color MediumVioletRed;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF191970.
	/// </summary>
	static const color MidnightBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF5FFFA.
	/// </summary>
	static const color MintCream;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFE4E1.
	/// </summary>
	static const color MistyRose;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFE4B5.
	/// </summary>
	static const color Moccasin;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFDEAD.
	/// </summary>
	static const color NavajoWhite;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0080.
	/// </summary>
	static const color Navy;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFDF5E6.
	/// </summary>
	static const color OldLace;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF80800.
	/// </summary>
	static const color Olive;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF6B8E23.
	/// </summary>
	static const color OliveDrab;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFA50.
	/// </summary>
	static const color Orange;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF450.
	/// </summary>
	static const color OrangeRed;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDA70D6.
	/// </summary>
	static const color Orchid;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFEEE8AA.
	/// </summary>
	static const color PaleGoldenrod;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF98FB98.
	/// </summary>
	static const color PaleGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFAFEEEE.
	/// </summary>
	static const color PaleTurquoise;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDB7093.
	/// </summary>
	static const color PaleVioletRed;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFEFD5.
	/// </summary>
	static const color PapayaWhip;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFDAB9.
	/// </summary>
	static const color PeachPuff;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFCD853F.
	/// </summary>
	static const color Peru;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFC0CB.
	/// </summary>
	static const color Pink;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFDDA0DD.
	/// </summary>
	static const color Plum;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFB0E0E6.
	/// </summary>
	static const color PowderBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF80080.
	/// </summary>
	static const color Purple;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF00.
	/// </summary>
	static const color Red;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFBC8F8F.
	/// </summary>
	static const color RosyBrown;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF4169E1.
	/// </summary>
	static const color RoyalBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF8B4513.
	/// </summary>
	static const color SaddleBrown;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFA8072.
	/// </summary>
	static const color Salmon;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF4A460.
	/// </summary>
	static const color SandyBrown;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF2E8B57.
	/// </summary>
	static const color SeaGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFF5EE.
	/// </summary>
	static const color SeaShell;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFA0522D.
	/// </summary>
	static const color Sienna;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFC0C0C0.
	/// </summary>
	static const color Silver;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF87CEEB.
	/// </summary>
	static const color SkyBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF6A5ACD.
	/// </summary>
	static const color SlateBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF708090.
	/// </summary>
	static const color SlateGray;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFAFA.
	/// </summary>
	static const color Snow;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF0FF7F.
	/// </summary>
	static const color SpringGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF4682B4.
	/// </summary>
	static const color SteelBlue;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFD2B48C.
	/// </summary>
	static const color Tan;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF08080.
	/// </summary>
	static const color Teal;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFD8BFD8.
	/// </summary>
	static const color Thistle;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFF6347.
	/// </summary>
	static const color Tomato;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF40E0D0.
	/// </summary>
	static const color Turquoise;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFEE82EE.
	/// </summary>
	static const color Violet;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF5DEB3.
	/// </summary>
	static const color Wheat;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFFFF.
	/// </summary>
	static const color White;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFF5F5F5.
	/// </summary>
	static const color WhiteSmoke;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FFFFFF0.
	/// </summary>
	static const color Yellow;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF9ACD32.
	/// </summary>
	static const color YellowGreen;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #00000000.
	/// </summary>
	static const color Transparent;
	/// <summary>
	/// gets a system-defined color that has an ARGB value of #FF9ACD32.
	/// </summary>
	static const color Empty;

	static color from_argb(uint32_t argb);
	static color from_argb(uint8_t alpha, color base);
	static color from_argb(uint8_t r, uint8_t g, uint8_t b);
	static color from_argb(uint8_t a, uint8_t r, uint8_t g, uint8_t b);
	static color from_known_color(known_color color);
	static color from_name(const std::string &name);
	bool equals(color b) const;
	float get_brightness() const;
	float get_hue() const;
	float getsaturation() const;
	bool is_empty() const { return is_empty_; }
	bool is_known_color() const;
	bool is_named_color() const;
	bool is_system_color() const;
	uint32_t to_argb() const;
	known_color to_known_color() const;
	uint8_t get_alpha() const { return alpha_; }
	uint8_t get_red() const { return red_; }
	uint8_t get_green() const { return green_; }
	uint8_t get_blue() const { return blue_; }
	std::string get_name() const;

private:
	color();
	color(uint8_t a, uint8_t r, uint8_t g, uint8_t b);

	bool is_empty_;
	uint8_t alpha_;
	uint8_t red_;
	uint8_t green_;
	uint8_t blue_;
};

inline bool operator==(color a, color b)
{
	return a.equals(b);
}

inline bool operator!=(color a, color b)
{
	return a != b;
}

} // namespace xlnt
