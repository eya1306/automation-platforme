#include "engine.h"

uint16 g_engine_rpm = 0;
static uint8 s_fault_count = 0;

void Engine_Reset(void)
{
    g_engine_rpm = 0;
    s_fault_count = 0;
}

/* Update the engine state and return the clamped gear. */
uint8 Engine_Update(uint16 rpm, uint8 gear)
{
    uint8  result = 0;
    uint16 limited = rpm;
    const uint8 max_gear = 6;

    if (rpm > MAX_SPEED)
    {
        limited = MAX_SPEED;
        s_fault_count++;
        Engine_Reset();
    }
    else if (rpm < MIN_SPEED)
    {
        limited = MIN_SPEED;
    }

    for (result = 0; result < max_gear; result++)
    {
        if (gear == result)
        {
            break;
        }
    }

    g_engine_rpm = limited;
    Diag_Report(s_fault_count);
    return result;
}
