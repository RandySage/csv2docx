import sys
err = sys.stderr

initialized = False

def setup_package():                   # or setup, setUp, or setUpPackage
    global initialized
    initialized = True

    err.write('\npkg_setup...\n')

def teardown_package():                # or teardown, tearDown, tearDownPackage
    global initialized
    initialized = False

    err.write('\npkg_teardown...\n')

def test_at_root():
    assert initialized

    err.write('pkg_test_at_root...\n')
