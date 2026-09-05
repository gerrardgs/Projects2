import os
import json

def _trained_model_path():
    # Located at ../train/trainfile/trained_model.json relative to this file
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, '..', 'train', 'trainfile', 'trained_model.json')


def get_default_mf():
    path = _trained_model_path()
    if os.path.exists(path):
        try:
            with open(path, 'r') as f:
                model = json.load(f)
            mf_raw = model.get('mf_params')
            if isinstance(mf_raw, dict):
                # Ensure tuples/lists are preserved as (type, params)
                mf = {}
                for var_k, cats in mf_raw.items():
                    mf[var_k] = {}
                    for cat, val in cats.items():
                        try:
                            mf[var_k][cat] = (val[0], val[1])
                        except Exception:
                            mf[var_k][cat] = val
                return mf
        except Exception:
            # Fallthrough to defaults
            pass

    # Hard-coded sensible defaults (small, realistic ranges)
    defaults = {
        'usia': {
            'Muda':  ('trap', [0.0, 0.0, 6.0, 18.0]),
            'Sedang':('tri',  [12.0, 30.0, 48.0]),
            'Tua':   ('trap', [42.0, 52.0, 59.0, 59.0]),
        },
        'mpasi': {
            'TerlaCepat': ('trap', [0.0, 0.0, 3.0, 5.0]),
            'Normal':     ('trap', [4.0, 5.5, 6.5, 8.0]),
            'Lambat':     ('trap', [7.0, 8.0, 12.0, 12.0]),
        },
        # TB: laki / perempuan (cm) - conservative human-scale values
        'tb_l': {
            'Pendek': ('trap', [38.0, 38.0, 55.0, 70.0]),
            'Normal': ('trap', [60.0, 70.0, 100.0, 110.0]),
            'Tinggi': ('trap', [98.0, 108.0, 110.0, 110.0]),
        },
        'tb_p': {
            'Pendek': ('trap', [38, 38, 53, 68]),
            'Normal': ('trap', [58, 68, 98, 108]),
            'Tinggi': ('trap', [98, 108, 130, 130]),
        },
        # Berat badan (gram) - per JK; values are illustrative
        'bb_l': {
            'Kurang':   ('trap', [1500, 1500, 5000, 8000]),
            'Normal':   ('trap', [6000, 8000, 15000, 18000]),
            'Lebih':    ('tri',  [15000, 18000, 22000]),
            'Obesitas': ('trap', [20000, 23000, 30000, 30000]),
        },
        'bb_p': {
            'Kurang':   ('trap', [1500, 1500, 4500, 7500]),
            'Normal':   ('trap', [5500, 7500, 14000, 17000]),
            'Lebih':    ('tri',  [14000, 17000, 21000]),
            'Obesitas': ('trap', [19000, 22000, 28000, 28000]),
        },
    }
    return defaults
#OK