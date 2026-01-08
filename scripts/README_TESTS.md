# Test Scripts

This directory contains test scripts to help verify the functionality of the Controller Macro Runner commands.

## Test Scripts

### test_rapid_buttons.json
**Rapid button change test** - Tests immediate serial transmission.
- Alternating buttons every 1ms
- Rapid sequences with no gaps
- Ultra-fast combos (frame-perfect inputs)
- Mashing faster than keepalive (20ms, 5ms cycles)
- Back-to-back button presses
- **Duration**: ~5 seconds
- **Use case**: Verify no button presses are missed due to transmission delays

### test_timing_precision.json
**Timing precision test** - Tests the high-precision timing system.
- Sub-10ms wait precision
- Rapid sequential presses (1ms intervals)
- Ultra-fast mashing (50, 100, 200 presses/sec)
- Mixed fractional millisecond timings
- **Duration**: ~5 seconds
- **Use case**: Verify sub-millisecond timing accuracy

### test_quick.json
**Quick verification test** - Fast test to verify basic mash functionality.
- Single button mash
- Multiple buttons mash
- Fast and slow mashing speeds
- **Duration**: ~3 seconds
- **Use case**: Quick sanity check

### test_mash_basic.json
**Basic mash command test** - Tests fundamental mash command features.
- Default mash settings (20 presses/second)
- Different single buttons
- Multiple button combinations
- **Duration**: ~6 seconds
- **Use case**: Verify basic mash functionality

### test_mash_speeds.json
**Mash speed variations** - Tests different mashing speeds.
- Default speed (20 presses/sec)
- Fast speed (40 presses/sec)
- Very fast speed (50 presses/sec)
- Slow speed (10 presses/sec)
- Very slow speed (5 presses/sec)
- **Duration**: ~15 seconds
- **Use case**: Compare different mashing rates

### test_all_buttons.json
**All buttons test** - Tests mashing with all available buttons.
- Face buttons (A, B, X, Y)
- D-pad buttons (Up, Down, Left, Right)
- Shoulder buttons (L, R)
- Start/Select buttons
- Button combinations
- **Duration**: ~15 seconds
- **Use case**: Verify all buttons work with mash

### test_variables_mash.json
**Variables with mash** - Tests using variables to control mash parameters.
- Variable-controlled duration, hold_ms, wait_ms
- Loops with changing speeds
- Conditional mashing
- Counter-based control
- **Duration**: ~10 seconds
- **Use case**: Verify variable support and dynamic configuration

### test_comprehensive.json
**Comprehensive command test** - Tests multiple commands working together.
- Basic press/wait commands
- Hold and release
- Mashing with different configurations
- Variables and loops
- Conditional execution
- Label/goto patterns
- Command combinations and sequences
- **Duration**: ~20 seconds
- **Use case**: Full integration test

## Running Tests

1. Open Controller Macro Runner
2. Connect to your serial controller or 3DS
3. Select a test script from the dropdown menu
4. Click "Run" to execute the test
5. Click "Stop" to terminate early if needed

## Expected Results

All tests should:
- Execute without errors
- Send proper button commands to the controller
- Respect timing parameters
- Complete successfully
- Release all buttons when finished

## Customizing Tests

You can modify these test scripts to:
- Change button combinations
- Adjust mashing durations
- Test different speeds
- Add your own test scenarios

## Notes

- Ensure your controller is connected before running tests
- Some tests use loops and may take longer to complete
- You can stop any test at any time using the Stop button
- Variables are reset when each script starts
