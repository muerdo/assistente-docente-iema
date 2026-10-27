#!/bin/bash
# Wrapper para kill_switch.sh mantendo compatibilidade total
exec "$(dirname "$(readlink -f "$0")")/kill_switch.sh" "$@"
