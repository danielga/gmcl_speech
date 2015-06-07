#include <GarrysMod/Lua/Interface.h>
#include <lua.hpp>
#include <cstdint>
#include <sapi.h>
#include <sphelper.h>
#include <comdef.h>
#include <vector>
#include <locale>
#include <codecvt>
#include <string>

namespace tts
{

static const char *metaname = "ISpRecognizer";
static uint8_t metatype = 171;
static const char *tablename = "speech";
static const char *invalid_error = "invalid ISpRecognizer";

struct userdata
{
	ISpRecognizer *reco;
	uint8_t type;
	ISpRecoContext *context;
	ISpRecoGrammar *grammar;
};

inline std::wstring ConvertToUTF16( const char *str )
{
	std::wstring_convert< std::codecvt_utf8_utf16<wchar_t> > converter;
	return converter.from_bytes( str );
}

inline std::string ConvertToUTF8( const wchar_t *str )
{
	std::wstring_convert< std::codecvt_utf8_utf16<wchar_t> > converter;
	return converter.to_bytes( str );
}

static int32_t PushError( lua_State *state, const char *error, HRESULT hr )
{
	LUA->PushNil( );
	LUA->PushString( error );
	LUA->PushNumber( hr );
	return 3;
}

inline int32_t PushEvent( lua_State *state, const CSpEvent &ev )
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

	LUA->PushNumber( ev.wParam );
	LUA->SetField( -2, "wparam" );

	LUA->PushNumber( ev.lParam );
	LUA->SetField( -2, "lparam" );

	return 1;
}

inline void Check( lua_State *state, int32_t index )
{
	if( !LUA->IsType( index, metatype ) )
		luaL_typerror( state, index, metaname );
}

inline userdata *GetUserdata( lua_State *state, int32_t index )
{
	return static_cast<userdata *>( LUA->GetUserdata( index ) );
}

static ISpRecognizer *GetAndValidate( lua_State *state, int32_t index )
{
	Check( state, index );
	ISpRecognizer *reco = GetUserdata( state, index )->reco;
	if( reco == nullptr )
		LUA->ArgError( index, invalid_error );

	return reco;
}

static ISpRecoContext *GetContextAndValidate( lua_State *state, int32_t index )
{
	Check( state, index );
	ISpRecoContext *context = GetUserdata( state, index )->context;
	if( context == nullptr )
		LUA->ArgError( index, invalid_error );

	return context;
}

static userdata *Create( lua_State *state )
{
	userdata *udata = static_cast<userdata *>( LUA->NewUserdata( sizeof( userdata ) ) );
	udata->type = metatype;
	udata->reco = nullptr;
	udata->context = nullptr;
	udata->grammar = nullptr;

	LUA->CreateMetaTableType( metaname, metatype );
	LUA->SetMetaTable( -2 );

	LUA->CreateTable( );
	lua_setfenv( state, -2 );

	return udata;
}

inline int32_t SetRecognizer( lua_State *state, ISpRecognizer *reco, const char *name )
{
	std::wstring wide = L"Name=" + ConvertToUTF16( name );

	CComPtr<ISpObjectToken> bestreco;
	HRESULT hr = SpFindBestToken( SPCAT_RECOGNIZERS, wide.c_str( ), nullptr, &bestreco );
	if( FAILED( hr ) )
		return PushError( state, "find_failed", hr );

	hr = reco->SetRecognizer( bestreco );
	if( FAILED( hr ) )
		return PushError( state, "change_failed", hr );

	return 0;
}

LUA_FUNCTION_STATIC( gc )
{
	Check( state, 1 );
	userdata *udata = GetUserdata( state, 1 );

	ISpRecoGrammar *grammar = udata->grammar;
	if( grammar != nullptr )
	{
		grammar->Release( );
		udata->grammar = nullptr;
	}

	ISpRecoContext *context = udata->context;
	if( context != nullptr )
	{
		context->Release( );
		udata->context = nullptr;
	}

	ISpRecognizer *reco = udata->reco;
	if( reco != nullptr )
	{
		reco->Release( );
		udata->reco = nullptr;
	}

	return 0;
}

LUA_FUNCTION_STATIC( tostring )
{
	lua_pushfstring( state, "%s: %p", metaname, GetAndValidate( state, 1 ) );
	return 1;
}

LUA_FUNCTION_STATIC( eq )
{
	LUA->PushBool( GetAndValidate( state, 1 ) == GetAndValidate( state, 2 ) );
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

	lua_getfenv( state, 1 );
	LUA->Push( 2 );
	LUA->RawGet( -2 );
	return 1;
}

LUA_FUNCTION_STATIC( newindex )
{
	lua_getfenv( state, 1 );
	LUA->Push( 2 );
	LUA->Push( 3 );
	LUA->RawSet( -3 );
	return 0;
}

LUA_FUNCTION_STATIC( valid )
{
	Check( state, 1 );
	LUA->PushBool( GetUserdata( state, 1 )->reco != nullptr );
	return 1;
}

LUA_FUNCTION_STATIC( pause )
{
	ISpRecoContext *context = GetContextAndValidate( state, 1 );

	HRESULT hr = context->Pause( 0 );
	if( FAILED( hr ) )
		return PushError( state, "pause_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( resume )
{
	ISpRecoContext *context = GetContextAndValidate( state, 1 );

	HRESULT hr = context->Resume( 0 );
	if( FAILED( hr ) )
		return PushError( state, "resume_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( _state )
{
	ISpRecognizer *reco = GetAndValidate( state, 1 );

	HRESULT hr = S_OK;
	SPRECOSTATE recostate = SPRST_INACTIVE;
	if( LUA->Top( ) > 1 )
	{
		const char *s = LUA->CheckString( 2 );
		if( std::strcmp( s, "inactive" ) == 0 )
			recostate = SPRST_INACTIVE;
		else if( std::strcmp( s, "active" ) == 0 )
			recostate = SPRST_ACTIVE;
		else
			return PushError( state, "invalid_state", E_INVALIDARG );

		hr = reco->GetRecoState( &recostate );
		if( FAILED( hr ) )
			return PushError( state, "state_failed", hr );

		LUA->PushBool( true );
		return 1;
	}

	hr = reco->GetRecoState( &recostate );
	if( FAILED( hr ) )
		return PushError( state, "state_failed", hr );

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
			LUA->PushString( "inactive_purge" );
			break;

		default:
			LUA->PushString( "unknown" );
			break;
	}

	return 1;
}

LUA_FUNCTION_STATIC( interest )
{
	ISpRecoContext *context = GetContextAndValidate( state, 1 );
	ULONGLONG flags = SPFEI( static_cast<ULONGLONG>( LUA->CheckNumber( 2 ) ) );

	int32_t top = LUA->Top( );
	for( int32_t k = 3; k <= top; ++k )
		flags |= SPFEI( static_cast<ULONGLONG>( LUA->CheckNumber( k ) ) );

	HRESULT hr = context->SetInterest( flags, flags );
	if( FAILED( hr ) )
		return PushError( state, "interest_failed", hr );

	LUA->PushBool( true );
	return 1;
}

LUA_FUNCTION_STATIC( events )
{
	ISpRecoContext *context = GetContextAndValidate( state, 1 );
	ULONG num = 1;

	if( LUA->Top( ) > 1 )
		num = static_cast<ULONG>( LUA->CheckNumber( 2 ) );

	std::vector<CSpEvent> events( num );
	HRESULT hr = context->GetEvents( num, events.data( ), &num );
	if( FAILED( hr ) )
		return PushError( state, "retrieval_failed", hr );

	LUA->PushNumber( num );

	if( num == 0 )
		return 1;
	else if( num == 1 )
		PushEvent( state, events[0] );
	else
	{
		LUA->CreateTable( );

		for( ULONG k = 0; k < num; ++k )
		{
			LUA->PushNumber( k + 1 );
			PushEvent( state, events[k] );
			LUA->SetTable( -3 );
		}
	}

	return 2;
}

LUA_FUNCTION_STATIC( recognizer )
{
	ISpRecognizer *reco = GetAndValidate( state, 1 );
	if( LUA->Top( ) > 1 )
	{
		int32_t res = SetRecognizer( state, reco, LUA->CheckString( 2 ) );
		if( res != 0 )
			return res;

		LUA->PushBool( true );
		return 1;
	}

	CComPtr<ISpObjectToken> token;
	HRESULT hr = reco->GetRecognizer( &token );
	if( FAILED( hr ) )
		return PushError( state, "retrieval_failed", hr );

	LPWSTR name = nullptr;
	hr = token->GetStringValue( nullptr, &name );
	if( FAILED( hr ) )
		return PushError( state, "name_failed", hr );

	std::string narrow = ConvertToUTF8( name );

	LUA->PushString( narrow.c_str( ) );
	return 1;
}

LUA_FUNCTION_STATIC( create )
{
	const char *reconame = nullptr;
	if( LUA->Top( ) > 0 )
		reconame = LUA->CheckString( 1 );

	userdata *udata = Create( state );

	CComPtr<ISpRecognizer> reco;
	HRESULT hr = CoCreateInstance(
		CLSID_SpInprocRecognizer,
		NULL,
		CLSCTX_ALL,
		IID_ISpRecognizer,
		reinterpret_cast<void **>( &reco )
	);
	if( FAILED( hr ) )
		return PushError( state, "creation_failed", hr );

	CComPtr<ISpAudio> audio;
	hr = SpCreateDefaultObjectFromCategoryId( SPCAT_AUDIOIN, &audio );
	if( FAILED( hr ) )
		return PushError( state, "search_failed", hr );

	hr = reco->SetInput( audio, TRUE );
	if( FAILED( hr ) )
		return PushError( state, "input_failed", hr );

	CComPtr<ISpRecoContext> context;
	hr = reco->CreateRecoContext( &context );
	if( FAILED( hr ) )
		return PushError( state, "context_failed", hr );

	ULONGLONG interest = SPFEI( SPEI_RECOGNITION );
	hr = context->SetInterest( interest, interest );
	if( FAILED( hr ) )
		return PushError( state, "interest_failed", hr );

	hr = context->SetNotifyWin32Event( );
	if( FAILED( hr ) )
		return PushError( state, "notify_failed", hr );

	CComPtr<ISpRecoGrammar> grammar;
	hr = context->CreateGrammar( 0, &grammar );
	if( FAILED( hr ) )
		return PushError( state, "grammar_failed", hr );

	// TODO
	// This doesn't work with Microsoft Speech Platform SDK.
	// Since it's a server utility, it doesn't have grammars by default.
	// Add more object types like ISpRecoContext and ISpRecoGrammar to allow the developer
	// to interact with them directly?
	/*hr = grammar->LoadDictation( nullptr, SPLO_STATIC );
	if( FAILED( hr ) )
		return PushError( state, "dictation_failed", hr );

	hr = grammar->SetDictationState( SPRS_ACTIVE );
	if( FAILED( hr ) )
		return PushError( state, "state_failed", hr );*/

	if( reconame != nullptr )
	{
		int32_t res = SetRecognizer( state, reco, reconame );
		if( res != 0 )
			return res;
	}

	hr = reco->SetRecoState( SPRST_ACTIVE );
	if( FAILED( hr ) )
		return PushError( state, "activation_failed", hr );

	udata->reco = reco.Detach( );
	udata->context = context.Detach( );
	udata->grammar = grammar.Detach( );
	return 1;
}

LUA_FUNCTION_STATIC( get_recognizers )
{
	CComPtr<IEnumSpObjectTokens> recos;
	HRESULT hr = SpEnumTokens( SPCAT_RECOGNIZERS, nullptr, nullptr, &recos );
	if( FAILED( hr ) )
		return PushError( state, "retrieval_failed", hr );

	ULONG num = 0;
	hr = recos->GetCount( &num );
	if( FAILED( hr ) )
		return PushError( state, "count_failed", hr );

	LUA->CreateTable( );

	ULONG idx = 0;
	while( num-- != 0 )
	{
		CComPtr<ISpObjectToken> token;
		hr = recos->Next( 1, &token, nullptr );
		if( FAILED( hr ) )
			return PushError( state, "enum_failed", hr );

		LPWSTR value = nullptr;
		hr = token->GetStringValue( nullptr, &value );
		if( FAILED( hr ) )
			return PushError( state, "identifier_failed", hr );

		std::string narrow = ConvertToUTF8( value );

		LUA->PushNumber( ++idx );
		LUA->PushString( narrow.c_str( ) );
		LUA->SetTable( -3 );
	}

	return 1;
}

static void Initialize( lua_State *state )
{
	if( FAILED( CoInitialize( nullptr ) ) )
		LUA->ThrowError( "failed to initialize COM interfaces" );

	LUA->CreateMetaTableType( metaname, metatype );

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

	LUA->PushCFunction( _state );
	LUA->SetField( -2, "state" );

	LUA->PushCFunction( interest );
	LUA->SetField( -2, "interest" );

	LUA->PushCFunction( events );
	LUA->SetField( -2, "events" );

	LUA->PushCFunction( recognizer );
	LUA->SetField( -2, "recognizer" );

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

static void Deinitialize( lua_State *state )
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
	tts::Initialize( state );
	return 0;
}

GMOD_MODULE_CLOSE( )
{
	tts::Deinitialize( state );
	return 0;
}