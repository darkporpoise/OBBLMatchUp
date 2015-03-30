
; For those who would fear the license - don't. I tried to license it as liberal as possible.
; It really means you can do what ever you want with this.
; Donations are wellcome and will be accepted via PayPal address: trancexx at yahoo dot com
; Thank you for the shiny stuff :kiss:

#comments-start
	Copyright 2013 Dragana R. <trancexx at yahoo dot com>
	
	Licensed under the Apache License, Version 2.0 (the "License");
	you may not use this file except in compliance with the License.
	You may obtain a copy of the License at
	
	http://www.apache.org/licenses/LICENSE-2.0
	
	Unless required by applicable law or agreed to in writing, software
	distributed under the License is distributed on an "AS IS" BASIS,
	WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
	See the License for the specific language governing permissions and
	limitations under the License.
#comments-end

#include-once

Global Const $INTERNET_DEFAULT_PORT = 0
Global Const $INTERNET_DEFAULT_HTTP_PORT = 80
Global Const $INTERNET_DEFAULT_HTTPS_PORT = 443

Global Const $INTERNET_SCHEME_HTTP = 1
Global Const $INTERNET_SCHEME_HTTPS = 2

Global Const $ICU_ESCAPE = 0x80000000

; For WinHttpOpen
Global Const $WINHTTP_FLAG_ASYNC = 0x10000000

; For WinHttpOpenRequest  ;
Global Const $WINHTTP_FLAG_ESCAPE_PERCENT = 0x00000004
Global Const $WINHTTP_FLAG_NULL_CODEPAGE = 0x00000008
Global Const $WINHTTP_FLAG_ESCAPE_DISABLE = 0x00000040
Global Const $WINHTTP_FLAG_ESCAPE_DISABLE_QUERY = 0x00000080
Global Const $WINHTTP_FLAG_BYPASS_PROXY_CACHE = 0x00000100
Global Const $WINHTTP_FLAG_REFRESH = $WINHTTP_FLAG_BYPASS_PROXY_CACHE
Global Const $WINHTTP_FLAG_SECURE = 0x00800000

Global Const $WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0
Global Const $WINHTTP_ACCESS_TYPE_NO_PROXY = 1
Global Const $WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3

Global Const $WINHTTP_NO_PROXY_NAME = ""
Global Const $WINHTTP_NO_PROXY_BYPASS = ""

Global Const $WINHTTP_NO_REFERER = ""
Global Const $WINHTTP_DEFAULT_ACCEPT_TYPES = 0

Global Const $WINHTTP_NO_ADDITIONAL_HEADERS = ""
Global Const $WINHTTP_NO_REQUEST_DATA = ""

Global Const $WINHTTP_HEADER_NAME_BY_INDEX = ""
Global Const $WINHTTP_NO_OUTPUT_BUFFER = 0
Global Const $WINHTTP_NO_HEADER_INDEX = 0

Global Const $WINHTTP_ADDREQ_INDEX_MASK = 0x0000FFFF
Global Const $WINHTTP_ADDREQ_FLAGS_MASK = 0xFFFF0000
Global Const $WINHTTP_ADDREQ_FLAG_ADD_IF_NEW = 0x10000000
Global Const $WINHTTP_ADDREQ_FLAG_ADD = 0x20000000
Global Const $WINHTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = 0x40000000
Global Const $WINHTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = 0x01000000
Global Const $WINHTTP_ADDREQ_FLAG_COALESCE = $WINHTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA
Global Const $WINHTTP_ADDREQ_FLAG_REPLACE = 0x80000000

Global Const $WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH = 0

; For WinHttp{Set and Query} Options  ;
Global Const $WINHTTP_OPTION_CALLBACK = 1
Global Const $WINHTTP_FIRST_OPTION = $WINHTTP_OPTION_CALLBACK
Global Const $WINHTTP_OPTION_RESOLVE_TIMEOUT = 2
Global Const $WINHTTP_OPTION_CONNECT_TIMEOUT = 3
Global Const $WINHTTP_OPTION_CONNECT_RETRIES = 4
Global Const $WINHTTP_OPTION_SEND_TIMEOUT = 5
Global Const $WINHTTP_OPTION_RECEIVE_TIMEOUT = 6
Global Const $WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT = 7
Global Const $WINHTTP_OPTION_HANDLE_TYPE = 9
Global Const $WINHTTP_OPTION_READ_BUFFER_SIZE = 12
Global Const $WINHTTP_OPTION_WRITE_BUFFER_SIZE = 13
Global Const $WINHTTP_OPTION_PARENT_HANDLE = 21
Global Const $WINHTTP_OPTION_EXTENDED_ERROR = 24
Global Const $WINHTTP_OPTION_SECURITY_FLAGS = 31
Global Const $WINHTTP_OPTION_SECURITY_CERTIFICATE_STRUCT = 32
Global Const $WINHTTP_OPTION_URL = 34
Global Const $WINHTTP_OPTION_SECURITY_KEY_BITNESS = 36
Global Const $WINHTTP_OPTION_PROXY = 38
Global Const $WINHTTP_OPTION_USER_AGENT = 41
Global Const $WINHTTP_OPTION_CONTEXT_VALUE = 45
Global Const $WINHTTP_OPTION_CLIENT_CERT_CONTEXT = 47
Global Const $WINHTTP_OPTION_REQUEST_PRIORITY = 58
Global Const $WINHTTP_OPTION_HTTP_VERSION = 59
Global Const $WINHTTP_OPTION_DISABLE_FEATURE = 63
Global Const $WINHTTP_OPTION_CODEPAGE = 68
Global Const $WINHTTP_OPTION_MAX_CONNS_PER_SERVER = 73
Global Const $WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER = 74
Global Const $WINHTTP_OPTION_AUTOLOGON_POLICY = 77
Global Const $WINHTTP_OPTION_SERVER_CERT_CONTEXT = 78
Global Const $WINHTTP_OPTION_ENABLE_FEATURE = 79
Global Const $WINHTTP_OPTION_WORKER_THREAD_COUNT = 80
Global Const $WINHTTP_OPTION_PASSPORT_COBRANDING_TEXT = 81
Global Const $WINHTTP_OPTION_PASSPORT_COBRANDING_URL = 82
Global Const $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH = 83
Global Const $WINHTTP_OPTION_SECURE_PROTOCOLS = 84
Global Const $WINHTTP_OPTION_ENABLETRACING = 85
Global Const $WINHTTP_OPTION_PASSPORT_SIGN_OUT = 86
Global Const $WINHTTP_OPTION_PASSPORT_RETURN_URL = 87
Global Const $WINHTTP_OPTION_REDIRECT_POLICY = 88
Global Const $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS = 89
Global Const $WINHTTP_OPTION_MAX_HTTP_STATUS_CONTINUE = 90
Global Const $WINHTTP_OPTION_MAX_RESPONSE_HEADER_SIZE = 91
Global Const $WINHTTP_OPTION_MAX_RESPONSE_DRAIN_SIZE = 92
Global Const $WINHTTP_OPTION_CONNECTION_INFO = 93
Global Const $WINHTTP_OPTION_CLIENT_CERT_ISSUER_LIST = 94
Global Const $WINHTTP_OPTION_SPN = 96
Global Const $WINHTTP_OPTION_GLOBAL_PROXY_CREDS = 97
Global Const $WINHTTP_OPTION_GLOBAL_SERVER_CREDS = 98
Global Const $WINHTTP_OPTION_UNLOAD_NOTIFY_EVENT = 99
Global Const $WINHTTP_OPTION_REJECT_USERPWD_IN_URL = 100
Global Const $WINHTTP_OPTION_USE_GLOBAL_SERVER_CREDENTIALS = 101
Global Const $WINHTTP_OPTION_RECEIVE_PROXY_CONNECT_RESPONSE = 103
Global Const $WINHTTP_OPTION_IS_PROXY_CONNECT_RESPONSE = 104
Global Const $WINHTTP_OPTION_SERVER_SPN_USED = 106
Global Const $WINHTTP_OPTION_PROXY_SPN_USED = 107
Global Const $WINHTTP_OPTION_SERVER_CBT = 108
Global Const $WINHTTP_LAST_OPTION = $WINHTTP_OPTION_SERVER_CBT
Global Const $WINHTTP_OPTION_USERNAME = 0x1000
Global Const $WINHTTP_OPTION_PASSWORD = 0x1001
Global Const $WINHTTP_OPTION_PROXY_USERNAME = 0x1002
Global Const $WINHTTP_OPTION_PROXY_PASSWORD = 0x1003

Global Const $WINHTTP_CONNS_PER_SERVER_UNLIMITED = 0xFFFFFFFF

Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM = 0
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_LOW = 1
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_HIGH = 2
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_DEFAULT = $WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM

Global Const $WINHTTP_OPTION_REDIRECT_POLICY_NEVER = 0
Global Const $WINHTTP_OPTION_REDIRECT_POLICY_DISALLOW_HTTPS_TO_HTTP = 1
Global Const $WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS = 2
Global Const $WINHTTP_OPTION_REDIRECT_POLICY_LAST = $WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS
Global Const $WINHTTP_OPTION_REDIRECT_POLICY_DEFAULT = $WINHTTP_OPTION_REDIRECT_POLICY_DISALLOW_HTTPS_TO_HTTP

Global Const $WINHTTP_DISABLE_PASSPORT_AUTH = 0x00000000
Global Const $WINHTTP_ENABLE_PASSPORT_AUTH = 0x10000000
Global Const $WINHTTP_DISABLE_PASSPORT_KEYRING = 0x20000000
Global Const $WINHTTP_ENABLE_PASSPORT_KEYRING = 0x40000000

Global Const $WINHTTP_DISABLE_COOKIES = 0x00000001
Global Const $WINHTTP_DISABLE_REDIRECTS = 0x00000002
Global Const $WINHTTP_DISABLE_AUTHENTICATION = 0x00000004
Global Const $WINHTTP_DISABLE_KEEP_ALIVE = 0x00000008
Global Const $WINHTTP_ENABLE_SSL_REVOCATION = 0x00000001
Global Const $WINHTTP_ENABLE_SSL_REVERT_IMPERSONATION = 0x00000002
Global Const $WINHTTP_DISABLE_SPN_SERVER_PORT = 0x00000000
Global Const $WINHTTP_ENABLE_SPN_SERVER_PORT = 0x00000001
Global Const $WINHTTP_OPTION_SPN_MASK = $WINHTTP_ENABLE_SPN_SERVER_PORT

; WinHTTP error codes  ;
Global Const $WINHTTP_ERROR_BASE = 12000
Global Const $ERROR_WINHTTP_OUT_OF_HANDLES = 12001
Global Const $ERROR_WINHTTP_TIMEOUT = 12002
Global Const $ERROR_WINHTTP_INTERNAL_ERROR = 12004
Global Const $ERROR_WINHTTP_INVALID_URL = 12005
Global Const $ERROR_WINHTTP_UNRECOGNIZED_SCHEME = 12006
Global Const $ERROR_WINHTTP_NAME_NOT_RESOLVED = 12007
Global Const $ERROR_WINHTTP_INVALID_OPTION = 12009
Global Const $ERROR_WINHTTP_OPTION_NOT_SETTABLE = 12011
Global Const $ERROR_WINHTTP_SHUTDOWN = 12012
Global Const $ERROR_WINHTTP_LOGIN_FAILURE = 12015
Global Const $ERROR_WINHTTP_OPERATION_CANCELLED = 12017
Global Const $ERROR_WINHTTP_INCORRECT_HANDLE_TYPE = 12018
Global Const $ERROR_WINHTTP_INCORRECT_HANDLE_STATE = 12019
Global Const $ERROR_WINHTTP_CANNOT_CONNECT = 12029
Global Const $ERROR_WINHTTP_CONNECTION_ERROR = 12030
Global Const $ERROR_WINHTTP_RESEND_REQUEST = 12032
Global Const $ERROR_WINHTTP_SECURE_CERT_DATE_INVALID = 12037
Global Const $ERROR_WINHTTP_SECURE_CERT_CN_INVALID = 12038
Global Const $ERROR_WINHTTP_CLIENT_AUTH_CERT_NEEDED = 12044
Global Const $ERROR_WINHTTP_SECURE_INVALID_CA = 12045
Global Const $ERROR_WINHTTP_SECURE_CERT_REV_FAILED = 12057
Global Const $ERROR_WINHTTP_CANNOT_CALL_BEFORE_OPEN = 12100
Global Const $ERROR_WINHTTP_CANNOT_CALL_BEFORE_SEND = 12101
Global Const $ERROR_WINHTTP_CANNOT_CALL_AFTER_SEND = 12102
Global Const $ERROR_WINHTTP_CANNOT_CALL_AFTER_OPEN = 12103
Global Const $ERROR_WINHTTP_HEADER_NOT_FOUND = 12150
Global Const $ERROR_WINHTTP_INVALID_SERVER_RESPONSE = 12152
Global Const $ERROR_WINHTTP_INVALID_HEADER = 12153
Global Const $ERROR_WINHTTP_INVALID_QUERY_REQUEST = 12154
Global Const $ERROR_WINHTTP_HEADER_ALREADY_EXISTS = 12155
Global Const $ERROR_WINHTTP_REDIRECT_FAILED = 12156
Global Const $ERROR_WINHTTP_SECURE_CHANNEL_ERROR = 12157
Global Const $ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT = 12166
Global Const $ERROR_WINHTTP_UNABLE_TO_DOWNLOAD_SCRIPT = 12167
Global Const $ERROR_WINHTTP_SECURE_INVALID_CERT = 12169
Global Const $ERROR_WINHTTP_SECURE_CERT_REVOKED = 12170
Global Const $ERROR_WINHTTP_NOT_INITIALIZED = 12172
Global Const $ERROR_WINHTTP_SECURE_FAILURE = 12175
Global Const $ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR = 12178
Global Const $ERROR_WINHTTP_SECURE_CERT_WRONG_USAGE = 12179
Global Const $ERROR_WINHTTP_AUTODETECTION_FAILED = 12180
Global Const $ERROR_WINHTTP_HEADER_COUNT_EXCEEDED = 12181
Global Const $ERROR_WINHTTP_HEADER_SIZE_OVERFLOW = 12182
Global Const $ERROR_WINHTTP_CHUNKED_ENCODING_HEADER_SIZE_OVERFLOW = 12183
Global Const $ERROR_WINHTTP_RESPONSE_DRAIN_OVERFLOW = 12184
Global Const $ERROR_WINHTTP_CLIENT_CERT_NO_PRIVATE_KEY = 12185
Global Const $ERROR_WINHTTP_CLIENT_CERT_NO_ACCESS_PRIVATE_KEY = 12186
Global Const $WINHTTP_ERROR_LAST = 12186

; WinHttp status codes  ;
Global Const $HTTP_STATUS_CONTINUE = 100
Global Const $HTTP_STATUS_SWITCH_PROTOCOLS = 101
Global Const $HTTP_STATUS_OK = 200
Global Const $HTTP_STATUS_CREATED = 201
Global Const $HTTP_STATUS_ACCEPTED = 202
Global Const $HTTP_STATUS_PARTIAL = 203
Global Const $HTTP_STATUS_NO_CONTENT = 204
Global Const $HTTP_STATUS_RESET_CONTENT = 205
Global Const $HTTP_STATUS_PARTIAL_CONTENT = 206
Global Const $HTTP_STATUS_WEBDAV_MULTI_STATUS = 207
Global Const $HTTP_STATUS_AMBIGUOUS = 300
Global Const $HTTP_STATUS_MOVED = 301
Global Const $HTTP_STATUS_REDIRECT = 302
Global Const $HTTP_STATUS_REDIRECT_METHOD = 303
Global Const $HTTP_STATUS_NOT_MODIFIED = 304
Global Const $HTTP_STATUS_USE_PROXY = 305
Global Const $HTTP_STATUS_REDIRECT_KEEP_VERB = 307
Global Const $HTTP_STATUS_BAD_REQUEST = 400
Global Const $HTTP_STATUS_DENIED = 401
Global Const $HTTP_STATUS_PAYMENT_REQ = 402
Global Const $HTTP_STATUS_FORBIDDEN = 403
Global Const $HTTP_STATUS_NOT_FOUND = 404
Global Const $HTTP_STATUS_BAD_METHOD = 405
Global Const $HTTP_STATUS_NONE_ACCEPTABLE = 406
Global Const $HTTP_STATUS_PROXY_AUTH_REQ = 407
Global Const $HTTP_STATUS_REQUEST_TIMEOUT = 408
Global Const $HTTP_STATUS_CONFLICT = 409
Global Const $HTTP_STATUS_GONE = 410
Global Const $HTTP_STATUS_LENGTH_REQUIRED = 411
Global Const $HTTP_STATUS_PRECOND_FAILED = 412
Global Const $HTTP_STATUS_REQUEST_TOO_LARGE = 413
Global Const $HTTP_STATUS_URI_TOO_LONG = 414
Global Const $HTTP_STATUS_UNSUPPORTED_MEDIA = 415
Global Const $HTTP_STATUS_RETRY_WITH = 449
Global Const $HTTP_STATUS_SERVER_ERROR = 500
Global Const $HTTP_STATUS_NOT_SUPPORTED = 501
Global Const $HTTP_STATUS_BAD_GATEWAY = 502
Global Const $HTTP_STATUS_SERVICE_UNAVAIL = 503
Global Const $HTTP_STATUS_GATEWAY_TIMEOUT = 504
Global Const $HTTP_STATUS_VERSION_NOT_SUP = 505
Global Const $HTTP_STATUS_FIRST = $HTTP_STATUS_CONTINUE
Global Const $HTTP_STATUS_LAST = $HTTP_STATUS_VERSION_NOT_SUP

Global Const $SECURITY_FLAG_IGNORE_UNKNOWN_CA = 0x00000100
Global Const $SECURITY_FLAG_IGNORE_CERT_DATE_INVALID = 0x00002000
Global Const $SECURITY_FLAG_IGNORE_CERT_CN_INVALID = 0x00001000
Global Const $SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE = 0x00000200
Global Const $SECURITY_FLAG_SECURE = 0x00000001
Global Const $SECURITY_FLAG_STRENGTH_WEAK = 0x10000000
Global Const $SECURITY_FLAG_STRENGTH_MEDIUM = 0x40000000
Global Const $SECURITY_FLAG_STRENGTH_STRONG = 0x20000000

Global Const $ICU_NO_ENCODE = 0x20000000
Global Const $ICU_DECODE = 0x10000000
Global Const $ICU_NO_META = 0x08000000
Global Const $ICU_ENCODE_SPACES_ONLY = 0x04000000
Global Const $ICU_BROWSER_MODE = 0x02000000
Global Const $ICU_ENCODE_PERCENT = 0x00001000

; Query flags  ;
Global Const $WINHTTP_QUERY_MIME_VERSION = 0
Global Const $WINHTTP_QUERY_CONTENT_TYPE = 1
Global Const $WINHTTP_QUERY_CONTENT_TRANSFER_ENCODING = 2
Global Const $WINHTTP_QUERY_CONTENT_ID = 3
Global Const $WINHTTP_QUERY_CONTENT_DESCRIPTION = 4
Global Const $WINHTTP_QUERY_CONTENT_LENGTH = 5
Global Const $WINHTTP_QUERY_CONTENT_LANGUAGE = 6
Global Const $WINHTTP_QUERY_ALLOW = 7
Global Const $WINHTTP_QUERY_PUBLIC = 8
Global Const $WINHTTP_QUERY_DATE = 9
Global Const $WINHTTP_QUERY_EXPIRES = 10
Global Const $WINHTTP_QUERY_LAST_MODIFIED = 11
Global Const $WINHTTP_QUERY_MESSAGE_ID = 12
Global Const $WINHTTP_QUERY_URI = 13
Global Const $WINHTTP_QUERY_DERIVED_FROM = 14
Global Const $WINHTTP_QUERY_COST = 15
Global Const $WINHTTP_QUERY_LINK = 16
Global Const $WINHTTP_QUERY_PRAGMA = 17
Global Const $WINHTTP_QUERY_VERSION = 18
Global Const $WINHTTP_QUERY_STATUS_CODE = 19
Global Const $WINHTTP_QUERY_STATUS_TEXT = 20
Global Const $WINHTTP_QUERY_RAW_HEADERS = 21
Global Const $WINHTTP_QUERY_RAW_HEADERS_CRLF = 22
Global Const $WINHTTP_QUERY_CONNECTION = 23
Global Const $WINHTTP_QUERY_ACCEPT = 24
Global Const $WINHTTP_QUERY_ACCEPT_CHARSET = 25
Global Const $WINHTTP_QUERY_ACCEPT_ENCODING = 26
Global Const $WINHTTP_QUERY_ACCEPT_LANGUAGE = 27
Global Const $WINHTTP_QUERY_AUTHORIZATION = 28
Global Const $WINHTTP_QUERY_CONTENT_ENCODING = 29
Global Const $WINHTTP_QUERY_FORWARDED = 30
Global Const $WINHTTP_QUERY_FROM = 31
Global Const $WINHTTP_QUERY_IF_MODIFIED_SINCE = 32
Global Const $WINHTTP_QUERY_LOCATION = 33
Global Const $WINHTTP_QUERY_ORIG_URI = 34
Global Const $WINHTTP_QUERY_REFERER = 35
Global Const $WINHTTP_QUERY_RETRY_AFTER = 36
Global Const $WINHTTP_QUERY_SERVER = 37
Global Const $WINHTTP_QUERY_TITLE = 38
Global Const $WINHTTP_QUERY_USER_AGENT = 39
Global Const $WINHTTP_QUERY_WWW_AUTHENTICATE = 40
Global Const $WINHTTP_QUERY_PROXY_AUTHENTICATE = 41
Global Const $WINHTTP_QUERY_ACCEPT_RANGES = 42
Global Const $WINHTTP_QUERY_SET_COOKIE = 43
Global Const $WINHTTP_QUERY_COOKIE = 44
Global Const $WINHTTP_QUERY_REQUEST_METHOD = 45
Global Const $WINHTTP_QUERY_REFRESH = 46
Global Const $WINHTTP_QUERY_CONTENT_DISPOSITION = 47
Global Const $WINHTTP_QUERY_AGE = 48
Global Const $WINHTTP_QUERY_CACHE_CONTROL = 49
Global Const $WINHTTP_QUERY_CONTENT_BASE = 50
Global Const $WINHTTP_QUERY_CONTENT_LOCATION = 51
Global Const $WINHTTP_QUERY_CONTENT_MD5 = 52
Global Const $WINHTTP_QUERY_CONTENT_RANGE = 53
Global Const $WINHTTP_QUERY_ETAG = 54
Global Const $WINHTTP_QUERY_HOST = 55
Global Const $WINHTTP_QUERY_IF_MATCH = 56
Global Const $WINHTTP_QUERY_IF_NONE_MATCH = 57
Global Const $WINHTTP_QUERY_IF_RANGE = 58
Global Const $WINHTTP_QUERY_IF_UNMODIFIED_SINCE = 59
Global Const $WINHTTP_QUERY_MAX_FORWARDS = 60
Global Const $WINHTTP_QUERY_PROXY_AUTHORIZATION = 61
Global Const $WINHTTP_QUERY_RANGE = 62
Global Const $WINHTTP_QUERY_TRANSFER_ENCODING = 63
Global Const $WINHTTP_QUERY_UPGRADE = 64
Global Const $WINHTTP_QUERY_VARY = 65
Global Const $WINHTTP_QUERY_VIA = 66
Global Const $WINHTTP_QUERY_WARNING = 67
Global Const $WINHTTP_QUERY_EXPECT = 68
Global Const $WINHTTP_QUERY_PROXY_CONNECTION = 69
Global Const $WINHTTP_QUERY_UNLESS_MODIFIED_SINCE = 70
Global Const $WINHTTP_QUERY_PROXY_SUPPORT = 75
Global Const $WINHTTP_QUERY_AUTHENTICATION_INFO = 76
Global Const $WINHTTP_QUERY_PASSPORT_URLS = 77
Global Const $WINHTTP_QUERY_PASSPORT_CONFIG = 78
Global Const $WINHTTP_QUERY_MAX = 78
Global Const $WINHTTP_QUERY_CUSTOM = 65535
Global Const $WINHTTP_QUERY_FLAG_REQUEST_HEADERS = 0x80000000
Global Const $WINHTTP_QUERY_FLAG_SYSTEMTIME = 0x40000000
Global Const $WINHTTP_QUERY_FLAG_NUMBER = 0x20000000

; Callback options  ;
Global Const $WINHTTP_CALLBACK_STATUS_RESOLVING_NAME = 0x00000001
Global Const $WINHTTP_CALLBACK_STATUS_NAME_RESOLVED = 0x00000002
Global Const $WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER = 0x00000004
Global Const $WINHTTP_CALLBACK_STATUS_CONNECTED_TO_SERVER = 0x00000008
Global Const $WINHTTP_CALLBACK_STATUS_SENDING_REQUEST = 0x00000010
Global Const $WINHTTP_CALLBACK_STATUS_REQUEST_SENT = 0x00000020
Global Const $WINHTTP_CALLBACK_STATUS_RECEIVING_RESPONSE = 0x00000040
Global Const $WINHTTP_CALLBACK_STATUS_RESPONSE_RECEIVED = 0x00000080
Global Const $WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION = 0x00000100
Global Const $WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED = 0x00000200
Global Const $WINHTTP_CALLBACK_STATUS_HANDLE_CREATED = 0x00000400
Global Const $WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING = 0x00000800
Global Const $WINHTTP_CALLBACK_STATUS_DETECTING_PROXY = 0x00001000
Global Const $WINHTTP_CALLBACK_STATUS_REDIRECT = 0x00004000
Global Const $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE = 0x00008000
Global Const $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE = 0x00010000
Global Const $WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE = 0x00020000
Global Const $WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE = 0x00040000
Global Const $WINHTTP_CALLBACK_STATUS_READ_COMPLETE = 0x00080000
Global Const $WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE = 0x00100000
Global Const $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR = 0x00200000
Global Const $WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE = 0x00400000
Global Const $WINHTTP_CALLBACK_FLAG_RESOLVE_NAME = 0x00000003
Global Const $WINHTTP_CALLBACK_FLAG_CONNECT_TO_SERVER = 0x0000000C
Global Const $WINHTTP_CALLBACK_FLAG_SEND_REQUEST = 0x00000030
Global Const $WINHTTP_CALLBACK_FLAG_RECEIVE_RESPONSE = 0x000000C0
Global Const $WINHTTP_CALLBACK_FLAG_CLOSE_CONNECTION = 0x00000300
Global Const $WINHTTP_CALLBACK_FLAG_HANDLES = 0x00000C00
Global Const $WINHTTP_CALLBACK_FLAG_DETECTING_PROXY = $WINHTTP_CALLBACK_STATUS_DETECTING_PROXY
Global Const $WINHTTP_CALLBACK_FLAG_REDIRECT = $WINHTTP_CALLBACK_STATUS_REDIRECT
Global Const $WINHTTP_CALLBACK_FLAG_INTERMEDIATE_RESPONSE = $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE
Global Const $WINHTTP_CALLBACK_FLAG_SECURE_FAILURE = $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
Global Const $WINHTTP_CALLBACK_FLAG_SENDREQUEST_COMPLETE = $WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE
Global Const $WINHTTP_CALLBACK_FLAG_HEADERS_AVAILABLE = $WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
Global Const $WINHTTP_CALLBACK_FLAG_DATA_AVAILABLE = $WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE
Global Const $WINHTTP_CALLBACK_FLAG_READ_COMPLETE = $WINHTTP_CALLBACK_STATUS_READ_COMPLETE
Global Const $WINHTTP_CALLBACK_FLAG_WRITE_COMPLETE = $WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
Global Const $WINHTTP_CALLBACK_FLAG_REQUEST_ERROR = $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
Global Const $WINHTTP_CALLBACK_FLAG_ALL_COMPLETIONS = 0x007E0000
Global Const $WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS = 0xFFFFFFFF

Global Const $API_RECEIVE_RESPONSE = 1
Global Const $API_QUERY_DATA_AVAILABLE = 2
Global Const $API_READ_DATA = 3
Global Const $API_WRITE_DATA = 4
Global Const $API_SEND_REQUEST = 5

Global Const $WINHTTP_HANDLE_TYPE_SESSION = 1
Global Const $WINHTTP_HANDLE_TYPE_CONNECT = 2
Global Const $WINHTTP_HANDLE_TYPE_REQUEST = 3

Global Const $WINHTTP_CALLBACK_STATUS_FLAG_CERT_REV_FAILED = 0x00000001
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_INVALID_CERT = 0x00000002
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_CERT_REVOKED = 0x00000004
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_INVALID_CA = 0x00000008
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_CERT_CN_INVALID = 0x00000010
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_CERT_DATE_INVALID = 0x00000020
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_CERT_WRONG_USAGE = 0x00000040
Global Const $WINHTTP_CALLBACK_STATUS_FLAG_SECURITY_CHANNEL_ERROR = 0x80000000

Global Const $WINHTTP_AUTH_SCHEME_BASIC = 0x00000001
Global Const $WINHTTP_AUTH_SCHEME_NTLM = 0x00000002
Global Const $WINHTTP_AUTH_SCHEME_PASSPORT = 0x00000004
Global Const $WINHTTP_AUTH_SCHEME_DIGEST = 0x00000008
Global Const $WINHTTP_AUTH_SCHEME_NEGOTIATE = 0x00000010

Global Const $WINHTTP_AUTH_TARGET_SERVER = 0x00000000
Global Const $WINHTTP_AUTH_TARGET_PROXY = 0x00000001


Global Const $WINHTTP_AUTOPROXY_AUTO_DETECT = 0x00000001
Global Const $WINHTTP_AUTOPROXY_CONFIG_URL = 0x00000002
Global Const $WINHTTP_AUTOPROXY_RUN_INPROCESS = 0x00010000
Global Const $WINHTTP_AUTOPROXY_RUN_OUTPROCESS_ONLY = 0x00020000
Global Const $WINHTTP_AUTO_DETECT_TYPE_DHCP = 0x00000001
Global Const $WINHTTP_AUTO_DETECT_TYPE_DNS_A = 0x00000002