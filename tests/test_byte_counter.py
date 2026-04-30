from app.core.byte_counter import ByteCounter, ByteCounterConfig


def test_korean_only():
    bc = ByteCounter(ByteCounterConfig(newline_bytes=1))
    total, _, warn = bc.analyze("가나다")
    assert total == 9
    assert warn is False


def test_english_only():
    bc = ByteCounter(ByteCounterConfig(newline_bytes=1))
    total, _, _ = bc.analyze("ABC")
    assert total == 3


def test_mixed_digits_spaces_symbols():
    bc = ByteCounter(ByteCounterConfig(newline_bytes=1))
    total, _, _ = bc.analyze("A1 가! ")
    assert total == 8


def test_newline_behavior():
    bc1 = ByteCounter(ByteCounterConfig(newline_bytes=1))
    bc2 = ByteCounter(ByteCounterConfig(newline_bytes=2))
    t1, _, _ = bc1.analyze("가\n나")
    t2, _, _ = bc2.analyze("가\n나")
    assert t1 == 7
    assert t2 == 8


def test_non_bmp_warning():
    bc = ByteCounter(ByteCounterConfig(newline_bytes=1))
    _, _, warn = bc.analyze("테스트😀")
    assert warn is True
