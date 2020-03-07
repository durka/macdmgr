#!/usr/bin/env nix-shell
#!nix-shell -i python3 -p "python3.withPackages(ps: [ps.ortools ps.xlrd ps.pyyaml])"

import sys
import yaml

from macd import Macd, Availability

with open(sys.argv[1], 'r') as f:
    cfg = yaml.load(f, Loader=yaml.Loader)

slv = Macd()

slv.xls(cfg['xls'])

if cfg.get('config', {}):
    if 'mgrs_per_night' in cfg['config']:
        slv.mgrs_per_night = cfg['config']['mgrs_per_night']
    if 'min_nights' in cfg['config']:
        slv.min_nights = cfg['config']['min_nights']
    if 'max_nights' in cfg['config']:
        slv.max_nights = cfg['config']['max_nights']
    if 'non_consecutive' in cfg['config']:
        slv.non_consecutive = cfg['config']['non_consecutive']
    if 'availability_level' in cfg['config']:
        slv.availability_level = Availability[cfg['config']['availability_level']]

for (m1, m2) in cfg.get('together', []):
    slv.keep_together(m1, m2)
for (m1, m2) in cfg.get('apart', []):
    slv.keep_apart(m1, m2)
for (m, ns) in cfg.get('pin', {}).items():
    if not isinstance(ns, list):
        ns = [ns]

    for n in ns:
        slv.pin_to(m, n)

slv.set_prev(cfg.get('prev'))

if slv.solve():
    print('Found a solution!')
    print()
    slv.print_managers()
    print()
    print('Raw data:')
    print(slv.raw_data())
else:
    print('No solution with %s' % cfg['config'])

