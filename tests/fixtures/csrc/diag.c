#include "engine.h"

void Diag_Report(uint8 faults)
{
    if (faults > 0)
    {
        Diag_Log(faults);
    }
}

void Diag_Log(uint8 code)
{
    g_engine_rpm = code;
}
