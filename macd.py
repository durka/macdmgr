from ortools.sat.python import cp_model
import re
from datetime import datetime
import xlrd
from enum import Enum, unique

@unique
class Availability(Enum):
    NO = 0
    IFNEEDBE = 1
    YES = 2

@unique
class Feasibility(Enum):
    UNKNOWN = cp_model.UNKNOWN
    MODEL_INVALID = cp_model.MODEL_INVALID
    INFEASIBLE = cp_model.INFEASIBLE
    FEASIBLE = cp_model.FEASIBLE
    OPTIMAL = cp_model.OPTIMAL

class Macd(object):
    def __init__(self):
        self.doodle = []
        self.nights = []
        self.prev = None
        self.together = []
        self.apart = []
        self.pin = {}
        self.solver = None

    # configuration

    mgrs_per_night = 2 # number of managers to assign each night
    min_nights = 1 # minimum number of nights to assign to each manager
    max_nights = 2 # maximum number of nights to assign to each manager
    non_consecutive = 2 # minimum number of nights in between managing
    availability_level = Availability.YES # set to IFNEEDBE if there is no feasible schedule with YES

    # export XLS from doodle and pass in here
    def xls(self, filename):
        xls = xlrd.open_workbook(filename).sheets()[0]

        self.name = re.fullmatch(r'Poll "(.*)"', xls.cell(0, 0).value).groups(1)[0]
        months = [c.value for c in xls.row(3)[1:]]
        nights = [c.value for c in xls.row(4)[1:]]

        cur_month = ''
        for i in range(len(nights)):
            if months[i] != '':
                cur_month = months[i]
            self.nights.append(datetime.strftime(datetime.strptime('%s %s' % (nights[i], cur_month), '%a %d %B %Y'), '%-m/%-d'))

        avail_map = {
            '': Availability.NO,
            '(OK)': Availability.IFNEEDBE,
            'OK': Availability.YES,
        }
        self.doodle = []
        for i in range(5, xls.nrows - 1):
            self.doodle.append((xls.cell(i, 0).value, [avail_map[t.value] for t in xls.row(i)[1:]]))

        self.n_mgrs = len(self.doodle)
        self.n_nights = len(self.nights)
        self.all_mgrs = range(self.n_mgrs)
        self.all_nights = range(self.n_nights)
        self.mgr_lookup = dict((self.doodle[m][0], m) for m in range(len(self.doodle)))
        self.night_lookup = dict((self.nights[n], n) for n in range(len(self.nights)))

    # set this to raw data output from a previous run, to make a new
    # schedule as similar as possible using updated inputs
    def set_prev(self, shifts):
        self.prev = shifts

    # some ways to put a thumb on the scale

    # pairs of managers who always manage together
    def keep_together(self, m1, m2):
        if m1 not in self.mgr_lookup:
            raise Exception('Unknown manager %s' % m1)
        if m2 not in self.mgr_lookup:
            raise Exception('Unknown manager %s' % m2)

        self.together.append((m1, m2))

    # pairs of managers who can't be put together
    def keep_apart(self, m1, m2):
        if m1 not in self.mgr_lookup:
            raise Exception('Unknown manager %s' % m1)
        if m2 not in self.mgr_lookup:
            raise Exception('Unknown manager %s' % m2)

        self.apart.append((m1, m2))

    # pin a manager to a specific date
    # can be called multiple times with the same manager to pin them to multiple dates
    # this will auto-exempt them from the separation constraint if it conflicts
    def pin_to(self, m, n):
        if m not in self.mgr_lookup:
            raise Exception('Unknown manager %s' % m)
        if n not in self.night_lookup:
            raise Exception('Unknown night %s' % n)

        if m in self.pin:
            self.pin[m].append(n)
        else:
            self.pin[m] = [n]

    def solve(self):
        if not self.doodle:
            raise Exception('Load a Doodle first')

        model = cp_model.CpModel()

        # Variables

        # shifts[m][n]: manager M manages on night N
        shifts = {}
        for m in self.all_mgrs:
            shifts[m] = {}
            for n in self.all_nights:
                shifts[m][n] = model.NewBoolVar('shift_m%in%i' % (m, n))

        # Constraints 

        # number of managers per night
        for n in self.all_nights:
            model.Add(sum(shifts[m][n] for m in self.all_mgrs) == self.mgrs_per_night)

        # number of nights per manager
        for m in self.all_mgrs:
            model.Add(sum(shifts[m][n] for n in self.all_nights) >= self.min_nights)
            model.Add(sum(shifts[m][n] for n in self.all_nights) <= self.max_nights)

        # availability
        for m in self.all_mgrs:
            for n in self.all_nights:
                if self.doodle[m][1][n].value < self.availability_level.value:
                    model.Add(shifts[m][n] == 0)

        # special dispensation
        for n in self.all_nights:
            for (a, b) in self.together:
                model.Add(shifts[self.mgr_lookup[a]][n] == shifts[self.mgr_lookup[b]][n])
            for (a, b) in self.apart:
                model.Add(shifts[self.mgr_lookup[a]][n] + shifts[self.mgr_lookup[b]][n] <= 1)

        nonconsexc = []
        for (m, ns) in self.pin.items():
            ns = sorted(ns)
            if len(ns) > 1 and min([self.night_lookup[j] - self.night_lookup[i] for i,j in zip(ns[:-1], ns[1:])]) <= self.non_consecutive:
                nonconsexc.append(self.mgr_lookup[m])

            for n in ns:
                model.Add(shifts[self.mgr_lookup[m]][self.night_lookup[n]] == 1)

        # non-consecutivity
        for m in self.all_mgrs:
            if m not in nonconsexc:
                for n in range(0, self.n_nights - self.non_consecutive):
                    model.Add(sum(shifts[m][j] for j in range(n, n + self.non_consecutive + 1)) <= 1)

        # ifneedbe
        model.Minimize(sum(shifts[m][n] for m in self.all_mgrs for n in self.all_nights if self.doodle[m][1][n] == Availability.IFNEEDBE))

        # conservativity
        if self.prev is not None:
            model.Maximize(sum(shifts[m][n] * self.prev[m][n] for m in self.all_mgrs for n in self.all_nights))

        solver = cp_model.CpSolver()
        solver.parameters.linearization_level = 0
        status = solver.Solve(model)

        self.shifts = shifts
        self.solver = solver

        return (status in [cp_model.OPTIMAL, cp_model.FEASIBLE])

    def print_managers(self):
        if self.solver is None:
            return 

        print('Managers for each night:')
        for n in self.all_nights:
            mgrs = []
            for m in self.all_mgrs:
                if self.solver.Value(self.shifts[m][n]) == 1:
                    if self.doodle[m][1][n] == Availability.IFNEEDBE:
                        mgrs.append('%s(!)' % self.doodle[m][0])
                    else:
                        mgrs.append(self.doodle[m][0])
            print('%s: %s' % (self.nights[n], ' and '.join(mgrs)))

        print('\nNights for each manager:')
        for m in self.all_mgrs:
            nights = []
            for n in self.all_nights:
                if self.solver.Value(self.shifts[m][n]) == 1:
                    if self.doodle[m][1][n] == Availability.IFNEEDBE:
                        nights.append('%s(!)' % self.nights[n])
                    else:
                        nights.append(self.nights[n])

            print('%s: %s' % (self.doodle[m][0], ', '.join(nights)))

    def raw_data(self):
        return {m: {n: self.solver.Value(s) for n, s in sched.items()} for (m, sched) in self.shifts.items()}

