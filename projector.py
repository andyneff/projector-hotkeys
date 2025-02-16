import re
import base64
import binascii

import win32gui
import win32con

import obspython as obs

DEFAULT_MONITOR = 1

PROGRAM_NAME = "Program Output"
MULTIVIEW_NAME = "Multiview Output"
STARTUP_NAME = "su"
GROUP_NAME = "gp"

monitors = {}
startup_projectors = {}
hotkey_ids = {}


def script_description():
  return """
  <center><h2>Windowed Projector Hotkeys</h2></center>
  <p>Hotkeys will be added for the Program output, Multiview, and each currently existing scene. Choose the monitor to which each output will be projected when the hotkey is pressed.</p>
  <p>You can also choose to open a projector to a specific monitor on startup. If you use this option, you may need to disable the "Save projectors on exit" preference or there will be duplicate projectors.</p>
  <p><b>If new scenes are added, or if scene names change, this script will need to be reloaded.</b></p>
  """


def script_properties():
  # set up the controls for the Program Output
  p = obs.obs_properties_create()

  gp = obs.obs_properties_create()
  obs.obs_properties_add_group(p, f"windowed_{GROUP_NAME}", "Windowed Projector", obs.OBS_GROUP_NORMAL, gp)
  obs.obs_properties_add_int(gp, "windowed_monitor", "Windowed Project to monitor:", 1, 10, 1)
  obs.obs_properties_add_int(gp, "windowed_left", "Windowed Project left:", 0, 16383, 1)
  obs.obs_properties_add_int(gp, "windowed_top", "Windowed Project top:", 0, 16383, 1)
  obs.obs_properties_add_int(gp, "windowed_width", "Windowed Project width:", 0, 16383, 1)
  obs.obs_properties_add_int(gp, "windowed_height", "Windowed Project height:", 0, 16383, 1)

  gp = obs.obs_properties_create()
  obs.obs_properties_add_group(p, f"{PROGRAM_NAME}_{GROUP_NAME}", "Program Output", obs.OBS_GROUP_NORMAL, gp)
  obs.obs_properties_add_int(gp, "{PROGRAM_NAME}", "Project to monitor:", -1, 10, 1)
  obs.obs_properties_add_bool(gp, f"{PROGRAM_NAME}{STARTUP_NAME}", "Open on Startup")

  # set up the controls for the Multiview
  gp = obs.obs_properties_create()
  obs.obs_properties_add_group(p, f"{MULTIVIEW_NAME}_{GROUP_NAME}", "Multiview", obs.OBS_GROUP_NORMAL, gp)
  obs.obs_properties_add_int(gp, "{MULTIVIEW_NAME}", "Project to monitor:", -1, 10, 1)
  obs.obs_properties_add_bool(gp, f"{MULTIVIEW_NAME}{STARTUP_NAME}", "Open on Startup")

  # loop through each scene and create a property group and control for choosing the monitor and startup settings
  scenes = obs.obs_frontend_get_scene_names()
  if scenes:
    for scene in scenes:
      gp = obs.obs_properties_create()
      obs.obs_properties_add_group(p, f"scene_{scene}_{GROUP_NAME}", f'{scene} (scene)', obs.OBS_GROUP_NORMAL, gp)
      obs.obs_properties_add_int(gp, f"scene_{scene}", "Project to monitor:", -1, 10, 1)
      obs.obs_properties_add_bool(gp, f"scene_{scene}{STARTUP_NAME}", "Open on Startup")

  sources = obs.obs_enum_sources()
  if sources:
    for source in sources:
      source = obs.obs_source_get_name(source)
      gp = obs.obs_properties_create()
      obs.obs_properties_add_group(p, f"source_{source}_{GROUP_NAME}", f'{source} (source)', obs.OBS_GROUP_NORMAL, gp)
      obs.obs_properties_add_int(gp, f"source_{source}", "Project to monitor:", -1, 10, 1)
      obs.obs_properties_add_bool(gp, f"source_{source}{STARTUP_NAME}", "Open on Startup")
  return p


def script_update(settings):
  update_monitor_preferences(settings)


def script_load(settings):
  def load_callback(e):
    if e == obs.OBS_FRONTEND_EVENT_FINISHED_LOADING:
      update_monitor_preferences(settings)
      register_hotkeys(settings)
      open_startup_projectors(settings)
      obs.remove_current_callback()

  scenes = obs.obs_frontend_get_scene_names()
  sources = obs.obs_enum_sources()
  if len(scenes) == 0 or len(sources) == 0:
    # on obs startup, scripts are loaded before scenes are finished loading
    # register a callback to register the hotkeys and open startup projectors after scenes are available
    obs.obs_frontend_add_event_callback(load_callback)
  else:
    # this runs when the script is loaded or reloaded from the settings window
    update_monitor_preferences(settings)
    register_hotkeys(settings)


def script_save(settings):
  for output, hotkey_id in hotkey_ids.items():
    hotkey_save_array = obs.obs_hotkey_save(hotkey_id)
    obs.obs_data_set_array(settings, output_to_function_name(output),
                           hotkey_save_array)
    obs.obs_data_array_release(hotkey_save_array)

def update_monitor_preference(settings, output):
  monitor = obs.obs_data_get_int(settings, output)
  if monitor is None or monitor == 0:
    monitor = DEFAULT_MONITOR

  if monitor >= 0:
    # monitors are 0 indexed here, but 1-indexed in the OBS menus
    monitor = monitor - 1

  monitors[output] = monitor

  # set which projectors should open on start up
  startup_projectors[output] = obs.obs_data_get_bool(settings, f"{output}{STARTUP_NAME}")

# find the monitor preferences for each projector and store them
def update_monitor_preferences(settings):
  update_monitor_preference(settings, PROGRAM_NAME)
  update_monitor_preference(settings, MULTIVIEW_NAME)
  for output in obs.obs_frontend_get_scene_names():
    update_monitor_preference(settings, f"scene_{output}")

  for output in [obs.obs_source_get_name(source)
                 for source in obs.obs_enum_sources()]:
    update_monitor_preference(settings, f"source_{output}")

def make_geometry(settings):
  projector_left = obs.obs_data_get_int(settings, "windowed_left")
  projector_top = obs.obs_data_get_int(settings, "windowed_top")
  projector_width = obs.obs_data_get_int(settings, "windowed_width")
  projector_height = obs.obs_data_get_int(settings, "windowed_height")
  projector_monitor = obs.obs_data_get_int(settings, "windowed_monitor")

  # precomputed values
  _magic_number = '01d9d0cb'
  _major_version = 3
  _minor_version = 0
  # geometry of the widget relative to its parent including any window frame
  _frame_left = projector_left
  _frame_top = projector_top
  _frame_right = _frame_left + projector_width
  _frame_bottom = _frame_top + projector_height
  # This property holds the geometry of the widget as it will appear when shown as a normal (not maximized or full screen) top-level widget
  _normal_left = _frame_left
  _normal_top = _frame_top
  _normal_right = _frame_right
  _normal_bottom = _frame_bottom

  _screen_number = projector_monitor
  _is_maximized = 0
  _is_fullscreen = 0
  _screen_width = 0
  # This property holds the geometry of the widget relative to its parent and excluding the window frame
  _screen_geometry_left = _frame_left
  _screen_geometry_top = _frame_top
  _screen_geometry_right = _frame_right
  _screen_geometry_bottom = _frame_bottom

  # concat all values
  d2hs = lambda x: f'{x:0{4}X}' # 4 digit hexes
  d2hl = lambda x: f'{x:0{8}X}' # 8 digit hexes

  geometry_hex = f"\
{_magic_number}\
{d2hs(_major_version)}\
{d2hs(_minor_version)}\
{d2hl(_frame_left)}\
{d2hl(_frame_top)}\
{d2hl(_frame_right)}\
{d2hl(_frame_bottom)}\
{d2hl(_normal_left)}\
{d2hl(_normal_top)}\
{d2hl(_normal_right)}\
{d2hl(_normal_bottom)}\
{d2hs(_screen_number)}\
{d2hs(_is_maximized)}\
{d2hs(_is_fullscreen)}\
{d2hl(_screen_width)}\
{d2hl(_screen_geometry_left)}\
{d2hl(_screen_geometry_top)}\
{d2hl(_screen_geometry_right)}\
{d2hl(_screen_geometry_bottom)}"

  # Encoded in Base64 using Qt’s geometry encoding
  return base64.b64encode(binascii.unhexlify(geometry_hex)).decode()

def register_hotkey(settings, output):
  # Different rules for everything... *Sigh*
  if output == MULTIVIEW_NAME:
    if monitors[output] == -1:
      title = "Multiview (Windowed)"
    else:
      title = "Multiview (Fullscreen)"
    hotkey_title = "Open Projector for Multiview"
  elif output == PROGRAM_NAME:
    if monitors[output] == -1:
      title = "Windowed Projector (Program)"
    else:
      title = "Fullscreen Projector (Program)"
    hotkey_title = "Open Projector for Program Output"
  else:
    projector_type, name = output.split('_', 1)
    if monitors[output] == -1:
      title = f"Windowed Projector ({projector_type.capitalize()}) - {name}"
    else:
      title = f"Fullscreen Projector ({projector_type.capitalize()}) - {name}"
    hotkey_title = f"Open Projector for {projector_type} '{name}'"

  def hotkey_pressed(pressed, output=output, title=title):
    if not pressed:
      window_handle = win32gui.FindWindow(None, title)
      if window_handle:
        win32gui.PostMessage(window_handle, win32con.WM_CLOSE, 0, 0)
      return
    open_projector(output, settings)

  hotkey_ids[output] = obs.obs_hotkey_register_frontend(
      output_to_function_name(output),
      hotkey_title, hotkey_pressed)

  hotkey_save_array = obs.obs_data_get_array(settings,
                                             output_to_function_name(output))
  obs.obs_hotkey_load(hotkey_ids[output], hotkey_save_array)
  obs.obs_data_array_release(hotkey_save_array)

# register a hotkey to open a projector for each output
def register_hotkeys(settings):
  register_hotkey(settings, PROGRAM_NAME)
  register_hotkey(settings, MULTIVIEW_NAME)
  for output in obs.obs_frontend_get_scene_names():
    register_hotkey(settings, f"scene_{output}")

  for output in [obs.obs_source_get_name(source)
                 for source in obs.obs_enum_sources()]:
    register_hotkey(settings, f"source_{output}")

# open a full screen projector
def open_projector(output, settings):
  # set the default monitor if one was never set
  if monitors.get(output, None) is None:
    monitors[output] = DEFAULT_MONITOR

  # call the front end API to open the projector
  if output == PROGRAM_NAME:
    # set the projector type if this is not a normal scene
    obs.obs_frontend_open_projector("StudioProgram", monitors[output], "", output)
  elif output == MULTIVIEW_NAME:
    # set the projector type if this is not a normal scene
    obs.obs_frontend_open_projector("Multiview", monitors[output], "", output)
  else:
    projector_type, name = output.split('_', 1)
    g = make_geometry(settings)
    obs.obs_frontend_open_projector(projector_type.capitalize(),
                                    monitors[output], g, name)

# open startup projectors
def open_startup_projectors(settings):
  for output, open_on_startup in startup_projectors.items():
    if open_on_startup:
      open_projector(output, settings)


# remove special characters from scene names to make them usable as function names
def output_to_function_name(name):
  return "ofsp_" + re.sub(r'\W', '_', name)
