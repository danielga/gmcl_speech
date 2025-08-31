#include <GarrysMod/Lua/Interface.h>
#include <GarrysMod/FactoryLoader.hpp>
#include <lua.hpp>

#include <cstdint>
#include <vector>
#include <string>

#include <filesystem.h>

#include <comdef.h>
#include <sapi.h>

#pragma warning( push )
// Disable warning C4996: 'GetVersionExA': was declared deprecated (sphelper.h:1319)
// Disable warning 4805: '==': unsafe mix of type 'BOOL' and type 'bool' in operation (sphelper.h:1319)
#pragma warning( disable: 4996 4805 )
#include <sphelper.h>
#pragma warning( pop )

namespace tts
{

static const char metaname[] = "ISpRecognizer";
static int32_t metatype = GarrysMod::Lua::Type::None;
static const char tablename[] = "speech";
static const char invalid_error[] = "invalid ISpRecognizer";

static SourceSDK::FactoryLoader filesystem_loader( "filesystem_stdio" );
static IFileSystem *filesystem = nullptr;

struct Container
{
	CComPtr<ISpRecognizer> reco;
	CComPtr<ISpRecoContext> context;
	CComPtr<ISpRecoGrammar> grammar;
};

inline std::wstring ConvertToUTF16( const char *str )
{
	const int wchars_num = MultiByteToWideChar( CP_UTF8, 0, str, -1, nullptr, 0 );
	std::wstring wstr( wchars_num - 1, L'\0' );
	MultiByteToWideChar( CP_UTF8, 0, str, -1, wstr.data( ), wchars_num );
	return wstr;
}

inline std::string ConvertToUTF8( const wchar_t *wstr )
{
	BOOL used_default_char = false;
	const int chars_num = WideCharToMultiByte( CP_UTF8, 0, wstr, -1, nullptr, 0, '\0', &used_default_char );
	std::string str( chars_num - 1, '\0' );
	WideCharToMultiByte( CP_UTF8, 0, wstr, -1, str.data( ), chars_num, '\0', &used_default_char );
	return str;
}

template<size_t maxLenInChars> inline int ValidateAndFixPathName( GarrysMod::Lua::ILuaBase *LUA, const char *path, char (&fullpath)[maxLenInChars] )
{
	char tmpfixedpath[2048] = { 0 };
	V_FixupPathName( tmpfixedpath, sizeof( tmpfixedpath ), path );
	const std::string fixedpath = tmpfixedpath;

	if( V_IsAbsolutePath( fixedpath.c_str( ) ) || fixedpath == ".." || fixedpath.find( ".." CORRECT_PATH_SEPARATOR_S ) != fixedpath.npos )
		return PushError( LUA, "invalid_path", 0 );

	if( !filesystem->RelativePathToFullPath_safe( fixedpath.c_str( ), "DATA", fullpath ) )
		return PushError( LUA, "path_failed", 0 );

	return 0;
}

static int32_t PushError( GarrysMod::Lua::ILuaBase *LUA, const char *error, HRESULT hr )
{
	LUA->PushNil( );
	LUA->PushString( error );
	LUA->PushNumber( hr );
	return 3;
}

inline int32_t PushEvent( GarrysMod::Lua::ILuaBase *LUA, const CSpEvent &ev )
{
	LUA->CreateTable( );

	LUA->PushNumber( ev.eEventId );
	LUA->SetField( -2, "event_id" );

	LUA->PushNumber( ev.elParamType );
	LUA->SetField( -2, "param_type" );

	LUA->PushNumber( ev.ulStreamNum );
	LUA->SetField( -2, "stream_num" );

	LUA->PushNumber( static_cast<double>( ev.ullAudioStreamOffset ) );
	LUA->SetField( -2, "stream_offset" );

	LUA->PushNumber( static_cast<double>( ev.wParam ) );
	LUA->SetField( -2, "wparam" );

	LUA->PushNumber( static_cast<double>( ev.lParam ) );
	LUA->SetField( -2, "lparam" );

	if( ev.eEventId == SPEI_RECOGNITION )
	{
		auto result = reinterpret_cast<ISpRecoResult *>( ev.lParam );
		LPWSTR pwszText = nullptr;
		HRESULT hr = result->GetText( static_cast<ULONG>( SP_GETWHOLEPHRASE ), static_cast<ULONG>( SP_GETWHOLEPHRASE ), TRUE, &pwszText, nullptr );
		if( SUCCEEDED( hr ) )
		{
			const auto text = ConvertToUTF8( pwszText );
			CoTaskMemFree( pwszText );

			LUA->PushString( text.c_str( ) );
			LUA->SetField( -2, "text" );
		}
		else
		{
			LUA->PushString( "<failure>" );
			LUA->SetField( -2, "text" );
		}
	}

	return 1;
}

inline void Check( GarrysMod::Lua::ILuaBase *LUA, int32_t index )
{
	if( !LUA->IsType( index, metatype ) )
		luaL_typerror( LUA->GetState( ), index, metaname );
}

inline Container *GetUserdata( GarrysMod::Lua::ILuaBase *LUA, int32_t index )
{
	return LUA->GetUserType<Container>( index, metatype );
}

static ISpRecognizer *GetAndValidate( GarrysMod::Lua::ILuaBase *LUA, int32_t index )
{
	Check( LUA, index );
	ISpRecognizer *reco = GetUserdata( LUA, index )->reco;
	if( reco == nullptr )
		LUA->ArgError( index, invalid_error );

	return reco;
}

static ISpRecoContext *GetContextAndValidate( GarrysMod::Lua::ILuaBase *LUA, int32_t index )
{
	Check( LUA, index );
	ISpRecoContext *context = GetUserdata( LUA, index )->context;
	if( context == nullptr )
		LUA->ArgError( index, invalid_error );

	return context;
}

static ISpRecoGrammar *GetGrammarAndValidate( GarrysMod::Lua::ILuaBase *LUA, int32_t index )
{
	Check( LUA, index );
	ISpRecoGrammar *grammar = GetUserdata( LUA, index )->grammar;
	if( grammar == nullptr )
		LUA->ArgError( index, invalid_error );

	return grammar;
}

static Container *Create( GarrysMod::Lua::ILuaBase *LUA )
{
	Container *udata = LUA->NewUserType<Container>( metatype );

	LUA->PushMetaTable( metatype );
	LUA->SetMetaTable( -2 );

	LUA->CreateTable( );
	lua_setfenv( LUA->GetState( ), -2 );

	return udata;
}

inline int32_t SetRecognizer( GarrysMod::Lua::ILuaBase *LUA, ISpRecognizer *reco, const char *id )
{
	const auto wId = ConvertToUTF16( id );

	CComPtr<ISpObjectToken> recotoken;
	HRESULT hr = SpGetTokenFromId( wId.c_str( ), &recotoken );
	if( FAILED( hr ) )
		return PushError( LUA, "find_failed", hr );

	hr = reco->SetRecognizer( recotoken );
	if( FAILED( hr ) )
		return PushError( LUA, "change_failed", hr );

	return 0;
}

LUA_FUNCTION_STATIC( gc )
{
	Check( LUA, 1 );
	Container *udata = GetUserdata( LUA, 1 );

	udata->grammar.Release( );
	udata->context.Release( );
	udata->reco.Release( );

	return 0;
}

LUA_FUNCTION_STATIC( tostring )
{
	lua_pushfstring( LUA->GetState( ), "%s: %p", metaname, GetAndValidate( LUA, 1 ) );
	return 1;
}

LUA_FUNCTION_STATIC( eq )
{
	LUA->PushBool( GetAndValidate( LUA, 1 ) == GetAndValidate( LUA, 2 ) );
	return 1;
}

LUA_FUNCTION_STATIC( index )
{
	LUA->GetMetaTable( 1 );
	LUA->Push( 2 );
	LUA->RawGet( -2 );
	if( !LUA->IsType( -1, GarrysMod::Lua::Type::NIL ) )
		return 1;

	LUA->Pop( 2 );

	lua_getfenv( LUA->GetState( ), 1 );
	LUA->Push( 2 );
	LUA->RawGet( -2 );
	return 1;
}

LUA_FUNCTION_STATIC( newindex )
{
	lua_getfenv( LUA->GetState( ), 1 );
	LUA->Push( 2 );
	LUA->Push( 3 );
	LUA->RawSet( -3 );
	return 0;
}

LUA_FUNCTION_STATIC( valid )
{
	Check( LUA, 1 );
	LUA->PushBool( GetUserdata( LUA, 1 )->reco != nullptr );
	return 1;
}

LUA_FUNCTION_STATIC( pause )
{
	auto context = GetContextAndValidate( LUA, 1 );

	HRESULT hr = context->Pause( 0 );
	if( FAILED( hr ) )
		return PushError( LUA, "pause_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( resume )
{
	auto context = GetContextAndValidate( LUA, 1 );

	HRESULT hr = context->Resume( 0 );
	if( FAILED( hr ) )
		return PushError( LUA, "resume_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( reco_state )
{
	auto reco = GetAndValidate( LUA, 1 );

	SPRECOSTATE recostate = SPRST_INACTIVE;
	if( LUA->Top( ) > 1 )
	{
		const char *s = LUA->CheckString( 2 );
		if( std::strcmp( s, "inactive" ) == 0 )
			recostate = SPRST_INACTIVE;
		else if( std::strcmp( s, "active" ) == 0 )
			recostate = SPRST_ACTIVE;
		else if( std::strcmp( s, "active_always" ) == 0 )
			recostate = SPRST_ACTIVE_ALWAYS;
		else if( std::strcmp( s, "inactive_with_purge" ) == 0 )
			recostate = SPRST_INACTIVE_WITH_PURGE;
		else
			return PushError( LUA, "invalid_state", E_INVALIDARG );

		HRESULT hr = reco->SetRecoState( recostate );
		if( FAILED( hr ) )
			return PushError( LUA, "state_failed", hr );

		LUA->PushBool( true );
		return 1;
	}

	HRESULT hr = reco->GetRecoState( &recostate );
	if( FAILED( hr ) )
		return PushError( LUA, "state_failed", hr );

	switch( recostate )
	{
		case SPRST_INACTIVE:
			LUA->PushString( "inactive" );
			break;

		case SPRST_ACTIVE:
			LUA->PushString( "active" );
			break;

		case SPRST_ACTIVE_ALWAYS:
			LUA->PushString( "active_always" );
			break;

		case SPRST_INACTIVE_WITH_PURGE:
			LUA->PushString( "inactive_with_purge" );
			break;

		default:
			LUA->PushString( "unknown" );
			break;
	}

	return 1;
}

LUA_FUNCTION_STATIC( dictation_state )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );

	SPRULESTATE dictationstate = SPRS_INACTIVE;
	const char *s = LUA->CheckString( 2 );
	if( std::strcmp( s, "inactive" ) == 0 )
		dictationstate = SPRS_INACTIVE;
	else if( std::strcmp( s, "active" ) == 0 )
		dictationstate = SPRS_ACTIVE;
	else if( std::strcmp( s, "active_with_autopause" ) == 0 )
		dictationstate = SPRS_ACTIVE_WITH_AUTO_PAUSE;
	else if( std::strcmp( s, "active_user_delimited" ) == 0 )
		dictationstate = SPRS_ACTIVE_USER_DELIMITED;
	else
		return PushError( LUA, "invalid_state", E_INVALIDARG );

	HRESULT hr = grammar->SetDictationState( dictationstate );
	if( FAILED( hr ) )
		return PushError( LUA, "state_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( grammar_state )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );

	SPGRAMMARSTATE grammarstate = SPGS_DISABLED;
	if( LUA->Top( ) > 1 )
	{
		const char *s = LUA->CheckString( 2 );
		if( std::strcmp( s, "disabled" ) == 0 )
			grammarstate = SPGS_DISABLED;
		else if( std::strcmp( s, "enabled" ) == 0 )
			grammarstate = SPGS_ENABLED;
		else if( std::strcmp( s, "exclusive" ) == 0 )
			grammarstate = SPGS_EXCLUSIVE;
		else
			return PushError( LUA, "invalid_state", E_INVALIDARG );

		HRESULT hr = grammar->SetGrammarState( grammarstate );
		if( FAILED( hr ) )
			return PushError( LUA, "state_failed", hr );

		LUA->PushBool( true );
		return 1;
	}

	HRESULT hr = grammar->GetGrammarState( &grammarstate );
	if( FAILED( hr ) )
		return PushError( LUA, "state_failed", hr );

	switch( grammarstate )
	{
		case SPGS_DISABLED:
			LUA->PushString( "disabled" );
			break;

		case SPGS_ENABLED:
			LUA->PushString( "enabled" );
			break;

		case SPGS_EXCLUSIVE:
			LUA->PushString( "exclusive" );
			break;

		default:
			LUA->PushString( "unknown" );
			break;
	}

	return 1;
}

LUA_FUNCTION_STATIC( interest )
{
	ISpRecoContext *context = GetContextAndValidate( LUA, 1 );
	ULONGLONG flags = SPFEI( static_cast<ULONGLONG>( LUA->CheckNumber( 2 ) ) );

	int32_t top = LUA->Top( );
	for( int32_t k = 3; k <= top; ++k )
		flags |= SPFEI( static_cast<ULONGLONG>( LUA->CheckNumber( k ) ) );

	HRESULT hr = context->SetInterest( flags, flags );
	if( FAILED( hr ) )
		return PushError( LUA, "interest_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( events )
{
	auto context = GetContextAndValidate( LUA, 1 );
	ULONG req = 1;

	if( LUA->Top( ) > 1 )
		req = static_cast<ULONG>( LUA->CheckNumber( 2 ) );

	std::vector<CSpEvent> events( req );
	ULONG num = 0;
	HRESULT hr = context->GetEvents( req, events.data( ), &num );
	if( FAILED( hr ) )
		return PushError( LUA, "retrieval_failed", hr );

	LUA->PushNumber( num );

	if( num == 0 )
		return 1;

	if( req == 1 && num == 1 )
	{
		PushEvent( LUA, events[0] );
		return 1;
	}

	LUA->CreateTable( );

	for( ULONG k = 0; k < num; ++k )
	{
		LUA->PushNumber( static_cast<double>( k + 1 ) );
		PushEvent( LUA, events[k] );
		LUA->SetTable( -3 );
	}

	return 2;
}

LUA_FUNCTION_STATIC( recognizer )
{
	auto reco = GetAndValidate( LUA, 1 );
	if( LUA->Top( ) > 1 )
	{
		int32_t res = SetRecognizer( LUA, reco, LUA->CheckString( 2 ) );
		if( res != 0 )
			return res;

		LUA->PushBool( true );
		return 1;
	}

	CComPtr<ISpObjectToken> token;
	HRESULT hr = reco->GetRecognizer( &token );
	if( FAILED( hr ) )
		return PushError( LUA, "retrieval_failed", hr );

	LPWSTR name = nullptr;
	hr = token->GetStringValue( nullptr, &name );
	if( FAILED( hr ) )
		return PushError( LUA, "name_failed", hr );

	const auto narrow = ConvertToUTF8( name );
	LUA->PushString( narrow.c_str( ) );
	return 1;
}

LUA_FUNCTION_STATIC( grammar )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );
	auto path = LUA->CheckString( 2 );

	char fullpath[2048] = { 0 };
	const int res = ValidateAndFixPathName( LUA, path, fullpath );
	if( res != 0 )
		return res;

	const auto wPath = ConvertToUTF16( fullpath );
	HRESULT hr = grammar->LoadCmdFromFile( wPath.c_str( ), SPLO_STATIC );
	if( FAILED( hr ) )
		return PushError( LUA, "load_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( load_dictation )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );

	HRESULT hr = grammar->LoadDictation( nullptr, SPLO_STATIC );
	if( FAILED( hr ) )
		return PushError( LUA, "load_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( unload_dictation )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );

	HRESULT hr = grammar->UnloadDictation( );
	if( FAILED( hr ) )
		return PushError( LUA, "unload_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( rule_state )
{
	auto grammar = GetGrammarAndValidate( LUA, 1 );

	std::wstring wName;
	if( LUA->IsType( 2, GarrysMod::Lua::Type::STRING ) )
	{
		const char *name = LUA->CheckString( 2 );
		wName = ConvertToUTF16( name );
	}

	SPRULESTATE rulestate = SPRS_INACTIVE;
	const char *s = LUA->CheckString( 3 );
	if( std::strcmp( s, "inactive" ) == 0 )
		rulestate = SPRS_INACTIVE;
	else if( std::strcmp( s, "active" ) == 0 )
		rulestate = SPRS_ACTIVE;
	else if( std::strcmp( s, "active_with_autopause" ) == 0 )
		rulestate = SPRS_ACTIVE_WITH_AUTO_PAUSE;
	else if( std::strcmp( s, "active_user_delimited" ) == 0 )
		rulestate = SPRS_ACTIVE_USER_DELIMITED;
	else
		return PushError( LUA, "invalid_state", E_INVALIDARG );

	HRESULT hr = grammar->SetRuleState( wName.empty( ) ? nullptr : wName.c_str( ), nullptr, rulestate);
	if( FAILED( hr ) )
		return PushError( LUA, "state_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( create )
{
	const char *recoid = nullptr;
	if( LUA->Top( ) > 0 )
		recoid = LUA->CheckString( 1 );

	Container *udata = Create( LUA );

	CComPtr<ISpRecognizer> reco;
	HRESULT hr = CoCreateInstance(
		CLSID_SpInprocRecognizer,
		NULL,
		CLSCTX_ALL,
		IID_ISpRecognizer,
		reinterpret_cast<void **>( &reco )
	);
	if( FAILED( hr ) )
		return PushError( LUA, "creation_failed", hr );

	CComPtr<ISpAudio> audio;
	hr = SpCreateDefaultObjectFromCategoryId( SPCAT_AUDIOIN, &audio );
	if( FAILED( hr ) )
		return PushError( LUA, "search_failed", hr );

	hr = reco->SetInput( audio, TRUE );
	if( FAILED( hr ) )
		return PushError( LUA, "input_failed", hr );

	CComPtr<ISpRecoContext> context;
	hr = reco->CreateRecoContext( &context );
	if( FAILED( hr ) )
		return PushError( LUA, "context_failed", hr );

	hr = context->SetNotifyWin32Event( );
	if( FAILED( hr ) )
		return PushError( LUA, "notify_failed", hr );

	CComPtr<ISpRecoGrammar> grammar;
	hr = context->CreateGrammar( 0, &grammar );
	if( FAILED( hr ) )
		return PushError( LUA, "grammar_failed", hr );

	if( recoid != nullptr )
	{
		int32_t res = SetRecognizer( LUA, reco, recoid );
		if( res != 0 )
			return res;
	}

	udata->reco = std::move( reco );
	udata->context = std::move( context );
	udata->grammar = std::move( grammar );
	return 1;
}

static bool insert_attributes( GarrysMod::Lua::ILua *LUA, CComPtr<ISpObjectToken> &token )
{
	CComPtr<ISpDataKey> attributesKey;
	HRESULT hr = token->OpenKey( L"Attributes", &attributesKey );
	if( FAILED( hr ) )
		return false;

	LPWSTR wLang = nullptr;
	hr = attributesKey->GetStringValue( L"Language", &wLang );
	if( FAILED( hr ) )
		return false;

	LUA->CreateTable( );

	int &errnoRef = errno;
	auto beg = wLang;
	wchar_t *end = nullptr;
	std::vector<int64_t> languages;
	for( int64_t k = 0, lang = std::wcstoll( beg, &end, 10 ); beg != end; k++, beg = end, lang = std::wcstoll( beg, &end, 10 ) )
	{
		if( errnoRef != 0 )
			continue;

		LUA->PushNumber( static_cast<double>( k ) );
		LUA->PushNumber( static_cast<double>( lang ) );
		LUA->SetTable( -3 );

		if( *end == ';' )
			++end;
	}

	LUA->SetField( -2, "languages" );

	::CoTaskMemFree( wLang );

	return true;
}

LUA_FUNCTION_STATIC( get_recognizers )
{
	CComPtr<IEnumSpObjectTokens> recos;
	HRESULT hr = SpEnumTokens( SPCAT_RECOGNIZERS, nullptr, nullptr, &recos );
	if( FAILED( hr ) )
		return PushError( LUA, "retrieval_failed", hr );

	ULONG num = 0;
	hr = recos->GetCount( &num );
	if( FAILED( hr ) )
		return PushError( LUA, "count_failed", hr );

	LUA->CreateTable( );

	ULONG idx = 0;
	while( num-- != 0 )
	{
		CComPtr<ISpObjectToken> token;
		hr = recos->Next( 1, &token, nullptr );
		if( FAILED( hr ) )
			return PushError( LUA, "enum_failed", hr );

		LPWSTR wId = nullptr;
		hr = token->GetId( &wId );
		if( FAILED( hr ) )
			return PushError( LUA, "identifier_failed", hr );

		const auto id = ConvertToUTF8( wId );
		CoTaskMemFree( wId );

		LPWSTR wName = nullptr;
		hr = token->GetStringValue( nullptr, &wName );
		if( FAILED( hr ) )
			return PushError( LUA, "name_failed", hr );

		const auto name = ConvertToUTF8( wName );

		LUA->PushNumber( ++idx );

		LUA->CreateTable( );

		LUA->PushString( id.c_str( ) );
		LUA->SetField( -2, "id" );

		LUA->PushString( name.c_str( ) );
		LUA->SetField( -2, "name" );

		insert_attributes( LUA, token );

		LUA->SetTable( -3 );
	}

	return 1;
}

static void Initialize( GarrysMod::Lua::ILuaBase *LUA )
{
	if( FAILED( CoInitialize( nullptr ) ) )
		LUA->ThrowError( "failed to initialize COM interfaces" );

	filesystem = filesystem_loader.GetInterface<IFileSystem>( FILESYSTEM_INTERFACE_VERSION );
	if( filesystem == nullptr )
		LUA->ThrowError( "unable to initialize IFileSystem" );

	metatype = LUA->CreateMetaTable( metaname );

	LUA->PushCFunction( gc );
	LUA->SetField( -2, "__gc" );

	LUA->PushCFunction( tostring );
	LUA->SetField( -2, "__tostring" );

	LUA->PushCFunction( eq );
	LUA->SetField( -2, "__eq" );

	LUA->PushCFunction( index );
	LUA->SetField( -2, "__index" );

	LUA->PushCFunction( newindex );
	LUA->SetField( -2, "__newindex" );

	LUA->PushCFunction( valid );
	LUA->SetField( -2, "valid" );

	LUA->PushCFunction( pause );
	LUA->SetField( -2, "pause" );

	LUA->PushCFunction( resume );
	LUA->SetField( -2, "resume" );

	LUA->PushCFunction( reco_state );
	LUA->SetField( -2, "reco_state" );

	LUA->PushCFunction( grammar_state );
	LUA->SetField( -2, "grammar_state" );

	LUA->PushCFunction( dictation_state );
	LUA->SetField( -2, "dictation_state" );

	LUA->PushCFunction( rule_state );
	LUA->SetField( -2, "rule_state" );

	LUA->PushCFunction( interest );
	LUA->SetField( -2, "interest" );

	LUA->PushCFunction( events );
	LUA->SetField( -2, "events" );

	LUA->PushCFunction( recognizer );
	LUA->SetField( -2, "recognizer" );

	LUA->PushCFunction( grammar );
	LUA->SetField( -2, "grammar" );

	LUA->PushCFunction( load_dictation );
	LUA->SetField( -2, "load_dictation" );

	LUA->PushCFunction( unload_dictation );
	LUA->SetField( -2, "unload_dictation" );

	LUA->Pop( 1 );

	LUA->PushSpecial( GarrysMod::Lua::SPECIAL_GLOB );

	LUA->CreateTable( );

	LUA->PushCFunction( create );
	LUA->SetField( -2, "create" );

	LUA->PushCFunction( get_recognizers );
	LUA->SetField( -2, "recognizers" );

	LUA->SetField( -2, tablename );

	LUA->Pop( 1 );
}

static void Deinitialize( GarrysMod::Lua::ILuaBase *LUA )
{
	LUA->PushSpecial( GarrysMod::Lua::SPECIAL_GLOB );

	LUA->PushNil( );
	LUA->SetField( -2, tablename );

	LUA->Pop( 1 );

	LUA->PushSpecial( GarrysMod::Lua::SPECIAL_REG );

	LUA->PushNil( );
	LUA->SetField( -2, metaname );

	LUA->Pop( 1 );

	::CoUninitialize( );
}

}

GMOD_MODULE_OPEN( )
{
	tts::Initialize( LUA );
	return 0;
}

GMOD_MODULE_CLOSE( )
{
	tts::Deinitialize( LUA );
	return 0;
}
