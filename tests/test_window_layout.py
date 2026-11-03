from window_layout import parent_relative_geometry, parent_relative_size


def test_work_pages_use_a_roomy_parent_relative_footprint():
    assert parent_relative_size(
        720, 620, parent_width=1800, parent_height=1200,
        screen_width=1920, screen_height=1600,
    ) == (1440, 1056)


def test_work_pages_stay_inside_parent_and_visible_screen():
    width, height, x, y = parent_relative_geometry(
        1120, 800, parent_width=900, parent_height=700,
        parent_x=40, parent_y=50, screen_width=1000, screen_height=800,
    )
    assert width <= 900 and height <= 700
    assert 12 <= x <= 1000 - width - 12
    assert 12 <= y <= 800 - height - 12
