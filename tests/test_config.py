from keyword_bible import config

def test_config_has_nine_semrush_files():
    assert len(config.SEMRUSH_FILES) == 9
    assert config.HOME == "policybazaar.ae"
    assert len(config.COMPETITORS) == 8

def test_seed_lists_nonempty_and_disjoint_verticals():
    assert config.HEALTH_SEEDS and config.TRAVEL_SEEDS and config.HARD_EXCLUDE
    # a term shouldn't be both a core health and core travel seed
    assert config.HEALTH_SEEDS.isdisjoint(config.TRAVEL_SEEDS)

def test_weights_present():
    assert config.SEVERITY_WEIGHT["Missing"] == 1.0
    assert config.INTENT_MULTIPLIER["transactional"] == 1.2
