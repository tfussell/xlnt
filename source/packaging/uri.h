#pragma once

#include <string>
#include <vector>

namespace xlnt {

/// <summary>
/// Defines the kinds of uris for the uri::IsWellFormeduriString(std::tring, uri_kind) and several other uri methods.
/// </summary>
enum class uri_kind
{
	/// <summary>
	/// The uri is an absolute uri.
	/// </summary>
	Absolute,
	/// <summary>
	/// The uri is a relative uri.
	/// </summary>
	Relative,
	/// <summary>
	/// The kind of the uri is indeterminate.
	/// </summary>
	RelativeOrAbsolute
};

/// <summary>
/// Defines host name types for the uri.CheckHostName method.
/// </summary>
enum class uri_host_name_type
{
	/// <summary>
	/// The host is set, but the type cannot be determined.
	/// </summary>
	Basic,
	/// <summary>
	/// The host name is a domain name system (DNS) style host name.
	/// </summary>
	Dns,
	/// <summary>
	/// The host name is an Internet Protocol (IP) version 4 host address.
	/// </summary>
	IPv4,
	/// <summary>
	/// The host name is an Internet Protocol (IP) version 6 host address.
	/// </summary>
	IPv6,
	/// <summary>
	/// The type of the host name is not supplied.
	/// </summary>
	Unknown
};

/// <summary>
/// Specifies the parts of a uri.
/// </summary>
enum class uri_components
{
	/// <summary>
	/// The Scheme, UserInfo, Host, Port, LocalPath, Query, and Fragment data.
	/// </summary>
	Absoluteuri,
	/// <summary>
	/// The Fragment data.
	/// </summary>
	Fragment,
	/// <summary>
	/// The Host data.
	/// </summary>
	Host,
	/// <summary>
	/// The Host and Port data. If no port data is in the uri and a default port has been assigned to the Scheme, the default port is returned. If there is no default port, -1 is returned.
	/// </summary>
	HostAndPort,
	/// <summary>
	/// The Scheme, Host, Port, LocalPath, and Query data.
	/// </summary>
	HttpRequesturi,
	/// <summary>
	/// Specifies that the delimiter should be included.
	/// </summary>
	KeepDelimiter,
	/// <summary>
	/// The normalized form of the Host.
	/// </summary>
	NormalizedHost,
	/// <summary>
	/// The LocalPath data.
	/// </summary>
	Path,
	/// <summary>
	/// The LocalPath and Query data. Also see PathAndQuery.
	/// </summary>
	PathAndQuery,
	/// <summary>
	/// The Port data.
	/// </summary>
	Port,
	/// <summary>
	/// The Query data.
	/// </summary>
	Query,
	/// <summary>
	/// The Scheme data.
	/// </summary>
	Scheme,
	/// <summary>
	/// The Scheme, Host, and Port data.
	/// </summary>
	SchemeAndServer,
	/// <summary>
	/// The Port data. If no port data is in the uri and a default port has been assigned to the Scheme, the default port is returned. If there is no default port, -1 is returned.
	/// </summary>
	StrongPort,
	/// <summary>
	/// The UserInfo data.
	/// </summary>
	UserInfo
};

/// <summary>
/// Controls how URI information is escaped.
/// </summary>
enum class uri_format
{
	/// <summary>
	/// Characters that have a reserved meaning in the requested URI components remain escaped. All others are not escaped. See Remarks.
	/// </summary>
	SafeUnescaped,
	/// <summary>
	/// No escaping is performed.
	/// </summary>
	Unescaped,
	/// <summary>
	/// Escaping is performed according to the rules in RFC 2396.
	/// </summary>
	uriEscaped
};

/// <summary>
/// Specifies the culture, case, and sort rules to be used by certain overloads of the String.Compare and String._equals methods
/// </summary>
enum class string_comparison
{
	/// <summary>
	/// Compare strings using culture-sensitive sort rules and the current culture.
	/// </summary>
	CurrentCulture,
	/// <summary>
	/// Compare strings using culture-sensitive sort rules, the current culture, and ignoring the case of the strings being compared.
	/// </summary>
	CurrentCultureIgnoreCase,
	/// <summary>
	/// Compare strings using culture-sensitive sort rules and the invariant culture.
	/// </summary>
	InvariantCulture,
	/// <summary>
	/// Compare strings using culture-sensitive sort rules, the invariant culture, and ignoring the case of the strings being compared.
	/// </summary>
	InvariantCultureIgnoreCase,
	/// <summary>
	/// Compare strings using ordinal sort rules.
	/// </summary>
	Ordinal,
	/// <summary>
	/// Compare strings using ordinal sort rules and ignoring the case of the strings being compared.
	/// </summary>
	OrdinalIgnoreCase
};

/// <summary>
/// Defines the parts of a URI for the uri::get_LeftPart method.
/// </summary>
enum class uri_partial
{
	/// <summary>
	/// The scheme and authority segments of the URI.
	/// </summary>
	Authority,
	/// <summary>
	/// The scheme, authority, and path segments of the URI.
	/// </summary>
	Path,
	/// <summary>
	/// The scheme, authority, path, and query segments of the URI.
	/// </summary>
	Query,
	/// <summary>
	/// The scheme segment of the URI.
	/// </summary>
	Scheme
};

/// <summary>
/// Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.
/// </summary>
class uri
{
public:
	/// <summary>
	/// Determines whether the specified host name is a valid DNS name.
	/// </summary>
	static uri_host_name_type CheckHostName(const std::string &name);

	/// <summary>
	/// Determines whether the specified scheme name is valid.
	/// </summary>
	static bool CheckSchemeName(const std::string &scheme_name);

	/// <summary>
	/// Compares the specified parts of two URIs using the specified comparison rules.
	/// </summary>
	static int Compare(const uri &uri1, const uri &uri2, const uri_components &parts_to_compare,
		uri_format compare_format, string_comparison comparison_type);

	/// <summary>
	/// Converts a string to its escaped representation.
	/// </summary>
	static std::string EscapeDataString(const std::string &string_to_escape);

	/// <summary>
	/// Converts a URI string to its escaped representation.
	/// </summary>
	static std::string EscapeuriString(const std::string &string_to_escape);

	/// <summary>
	/// gets the decimal value of a hexadecimal digit.
	/// </summary>
	static int FromHex(char digit);

	/// <summary>
	/// Converts a specified character into its hexadecimal equivalent.
	/// </summary>
	static std::string HexEscape(char character);

	/// <summary>
	/// Converts a specified hexadecimal representation of a character to the character.
	/// </summary>
	static char HexUnescape(const std::string &pattern, int &index);

	/// <summary>
	/// Determines whether a specified character is a valid hexadecimal digit.
	/// </summary>
	static bool IsHexDigit(char character);

	/// <summary>
	/// Determines whether a character in a string is hexadecimal encoded.
	/// </summary>
	static bool IsHexEncoding(const std::string &pattern, int index);

	/// <summary>
	/// Indicates whether the string is well-formed by attempting to construct a URI with the string and ensures that the string does not require further escaping.
	/// </summary>
	static bool IsWellFormeduriString(const std::string &uri_string, uri_kind uri_kind);

	/// <summary>
	/// Creates a new uri using the specified String instance and a uri_kind.
	/// </summary>
	static bool TryCreate(const std::string &uri_string, uri_kind uri_kind, uri &result);

	/// <summary>
	/// Creates a new uri using the specified base and relative String instances.
	/// </summary>
	static bool TryCreate(const uri &base_uri, const std::string &relative_uri, uri &result);

	/// <summary>
	/// Creates a new uri using the specified base and relative uri instances.
	/// </summary>
	static bool TryCreate(const uri &base_uri, const uri &relative_uri, uri &result);

	/// <summary>
	/// Converts a string to its unescaped representation.
	/// </summary>
	static std::string UnescapeDataString(const std::string &string_to_unescape);

	/// <summary>
	/// Initializes a new instance of the uri class with the specified URI. This constructor allows you to specify if the URI string is a relative URI, absolute URI, or is indeterminate.
	/// </summary>
	uri(const std::string &uri_string, uri_kind uri_kind = uri_kind::Absolute);

	/// <summary>
	/// Initializes a new instance of the uri class based on the specified base URI and relative URI string.
	/// </summary>
	uri(const uri &base_uri, const std::string &relative_uri);

	/// <summary>
	/// Initializes a new instance of the uri class based on the combination of a specified base uri instance and a relative uri instance.
	/// </summary>
	uri(const uri &base_uri, const uri &relative_uri);

	/// <summary>
	/// 
	/// </summary>
	uri &operator=(const uri &);

	/// <summary>
	/// gets the absolute path of the URI.
	/// </summary>
	std::string get_AbsolutePath() const;

	/// <summary>
	/// gets the absolute URI.
	/// </summary>
	std::string get_Absoluteuri() const;

	/// <summary>
	/// gets the Domain Name System (DNS) host name or IP address and the port number for a server.
	/// </summary>
	std::string get_Authority() const;

	/// <summary>
	/// gets an unescaped host name that is safe to use for DNS resolution.
	/// </summary>
	std::string get_DnsSafeHost() const;

	/// <summary>
	/// gets the escaped URI fragment.
	/// </summary>
	std::string get_Fragment() const;

	/// <summary>
	/// gets the host component of this instance.
	/// </summary>
	std::string get_Host() const;
	
	/// <summary>
	/// gets the type of the host name specified in the URI.
	/// </summary>
	uri_host_name_type get_HostNameType() const;

	/// <summary>
	/// gets whether the uri instance is absolute.
	/// </summary>
	bool IsAbsoluteuri() const;

	/// <summary>
	/// gets whether the port value of the URI is the default for this scheme.
	/// </summary>
	bool IsDefaultPort() const;

	/// <summary>
	/// gets a value indicating whether the specified uri is a file URI.
	/// </summary>
	bool IsFile() const;

	/// <summary>
	/// gets whether the specified uri references the local host.
	/// </summary>
	bool IsLoopback() const;

	/// <summary>
	/// gets whether the specified uri is a universal naming convention (UNC) path.
	/// </summary>
	bool IsUnc() const;

	/// <summary>
	/// gets a local operating-system representation of a file name.
	/// </summary>
	std::string get_LocalPath() const;

	/// <summary>
	/// gets the original URI string that was passed to the uri constructor.
	/// </summary>
	std::string get_OriginalString() const;

	/// <summary>
	/// gets the AbsolutePath and Query properties separated by a question mark (?).
	/// </summary>
	std::string get_PathAndQuery() const;

	/// <summary>
	/// gets the port number of this URI.
	/// </summary>
	int get_Port() const;

	/// <summary>
	/// gets any query information included in the specified URI.
	/// </summary>
	std::string get_Query() const;

	/// <summary>
	/// gets the scheme name for this URI.
	/// </summary>
	std::string get_Scheme() const;

	/// <summary>
	/// gets an array containing the path segments that make up the specified URI.
	/// </summary>
	std::vector<std::string> get_Segments() const;

	/// <summary>
	/// Indicates that the URI string was completely escaped before the uri instance was created.
	/// </summary>
	bool get_UserEscaped() const;

	/// <summary>
	/// gets the user name, password, or other user-specific information associated with the specified URI.
	/// </summary>
	std::string get_UserInfo() const;

	/// <summary>
	/// Compares two uri instances for equality.
	/// </summary>
	bool Equals(const uri &comparand);

	/// <summary>
	/// gets the specified components of the current instance using the specified escaping for special characters.
	/// </summary>
	std::string get_Components(uri_components components, uri_format format);

	/// <summary>
	/// gets the specified portion of a uri instance.
	/// </summary>
	std::string get_LeftPart(uri_partial part);

	/// <summary>
	/// Determines whether the current uri instance is a base of the specified uri instance.
	/// </summary>
	bool IsBaseOf(const uri &uri);

	/// <summary>
	/// Indicates whether the string used to construct this uri was well-formed and is not required to be further escaped.
	/// </summary>
	bool IsWellFormedOriginalString();

	/// <summary>
	/// Determines the difference between two uri instances.
	/// </summary>
	uri MakeRelativeuri(const uri &uri);

private:
	void Parse();

	void ParseAbsolute();

	void ParseRelative();

	void ParseAuthority();

	void ThrowIfNotAbsolute() const;

	std::string original_string_;

	uri_kind kind_;

	std::string absolute_path_;

	std::string absolute_uri_;

	std::string authority_;

	std::string dns_safe_host_;

	std::string fragment_;

	std::string host_;

	uri_host_name_type host_name_type_;

	bool is_default_port_;

	bool is_file_;

	bool is_unc_;

	bool is_loopback_;

	std::string local_path_;

	std::string path_and_query_;

	int port_;

	std::string query_;

	std::string scheme_;

	std::vector<std::string> segments_;

	bool user_escaped_;

	std::string user_info_;
};

} // namespace xlnt
