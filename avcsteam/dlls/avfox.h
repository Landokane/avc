int FoxCreate( edict_t *pEntity );
void FoxPrecache();
void FoxThink ( edict_t *pent );
void FoxTouch(edict_t *pent, edict_t *pentTouch);
void FoxTouch2(edict_t *pent, edict_t *pentTouch);
void AddAvPlace ( edict_t *pEntity );
void FoxSetSequence(int seq, entvars_t *pev);
void FoxSee( edict_t *pEntity );
void FoxStopSee( edict_t *pEntity );
void FoxControl( int button	);
void FoxStartControl( edict_t *pEntity );
void FoxStopControl( edict_t *pEntity );