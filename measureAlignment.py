# -*- coding: utf-8 -*-
import comtypes.client
from comtypes import COMError
from collections import OrderedDict
from input import ZERO, TOO_CLOSE, POINT_MANDATORY_STATIONS, \
                  POINTS_AT_GEOM_STATIONS, POINTS_AT_PVI_STATIONS, \
                  STARTING_STATION, ENDING_STATION, OFFSETS, \
                  STEP, TOLERANCE

# TODO: Implement program behavior for alignments without profile data

# -------------------------- Sample input.py --------------------------
# ZERO = 1e-5
# TOO_CLOSE = 0.10  # unit: meters
# POINT_MANDATORY_STATIONS = [409.60, 594.13, 854.47]  # list of stations set manually (for example shaft stations)
# POINTS_AT_GEOM_STATIONS = True
# POINTS_AT_PVI_STATIONS = True
# STARTING_STATION = 0.14
# ENDING_STATION = 925.58
# OFFSETS = [0.0, -2.0]  # unit: meters
# STEP = 10.0  # unit: meters
# TOLERANCE = 1.5  # unit: meters
# ------------------------- End sample input.py ------------------------

# ------------- Necessary Typelibs (AutoCAD Civil 3D 2008) -------------
# AutoCAD 2008 Type Library
#     CLSID:      "_851A4561_F4EC_4631_9B0C_E7DC407512C9_0_1_0"
#     DLL/TLB:    "C:\\Program Files\\Common Files\\Autodesk Shared\\acax17enu.tlb"
#
# AutoCAD/ObjectDBX Common 17.0 Type Library
#     CLSID:      "_11A32D00_9E89_4C16_82CB_629DEBA56AE2_0_1_0"
#     DLL/TLB:    "C:\\Program Files\\Common Files\\Autodesk Shared\\axdb17enu.tlb"
#
# AEC Base 5.5 Object Library
#     CLSID:      "_E7BAE300_BF69_431B_9153_1FF0DEF8F613_0_5_5"
#     DLL/TLB:    "C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXBase.tlb"
#
# AEC Base 5.5 Application Library
#     CLSID:      "_E7BAE100_BF69_431B_9153_1FF0DEF8F613_0_5_5"
#     DLL/TLB:    "C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXUIBase.tlb"
#
# Autodesk Civil Engineering Land 5.0 Object Library
#     CLSID:      "_E932008F_0032_4186_8DF9_3D5372B7DC57_0_5_0"
#     DLL/TLB:    "C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXLand.tlb"
#
# Autodesk Civil Engineering UI Land 5.0 Object Library
#     CLSID:      "_E9330001_0032_4186_8DF9_3D5372B7DC57_0_5_0"
#     DLL/TLB:    "C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXUiLand.tlb"
#
# OLE Automation (2.0)
#     CLSID:      "_00020430_0000_0000_C000_000000000046_0_2_0"
#     DLL/TLB:    "C:\\WINDOWS\\system32\\stdole2.tlb"
# ----------- End Necessary Typelibs (AutoCAD Civil 3D 2008) -----------

# # Generate modules of necessary typelibs for comtypes package
# # Missing typelibs from above table are generated automatically
# comtypes.client.GetModule("C:\\Program Files\\Common Files\\Autodesk Shared\\acax17enu.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXUIBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXLand.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXUiLand.tlb")
# raise SystemExit

TLB = comtypes.client.GetModule(
    "C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXLand.tlb")

# Get running instance of the AutoCAD application
acadApp = comtypes.client.GetActiveObject("AutoCAD.Application")
aeccApp = acadApp.GetInterfaceObject("AeccXUiLand.AeccApplication.5.0")

# Get the Document object and alignment
doc = aeccApp.ActiveDocument
alignment, point_clicked = doc.Utility.GetEntity("Select an alignment:")


def isalmostzero(num, zero=ZERO):
    """Return True if givven number <num> is equal or less than a certain
    value <zero>. <zero> should be small enough to be considered insignificant
    in the context it refers to.
    """
    return abs(num) <= zero


def isalmostequal(num1, num2, zero=ZERO):
    """Return True if <num1> and <num2> are considered equal "enough" in a
    certain context. Use this function instead of direct equality comparison
    of float numbers.
    """
    return isalmostzero(num1 - num2, zero)


def isnuminiterable(num, iterable, zero=ZERO):
    """Return True if <iterable> contains a number that would be considered
    equal "enough" to <num> in a certain context.
    """
    for elem in iterable:
        if isalmostequal(num, elem, zero):
            return True
    return False


def issuewarning(msg):
    print 50 * "!"
    print msg
    print 50 * "!"

# Prepare a list of stations where points are going to be created
if STARTING_STATION is None:
    STARTING_STATION = alignment.StartingStation
if ENDING_STATION is None:
    ENDING_STATION = alignment.EndingStation
assert not isalmostequal(STARTING_STATION, ENDING_STATION, TOO_CLOSE)
pointStations = [STARTING_STATION, ENDING_STATION]

# Get Alignment entities if needed
if POINTS_AT_GEOM_STATIONS:
    entities = {}
    for ent in alignment.Entities:
        if ent.Type in (TLB.aeccTangent, TLB.aeccArc, TLB.aeccSpiral):
            entities[ent.StartingStation] = ent
        elif ent.Type == TLB.aeccSpiralCurveSpiralGroup:
            entities[ent.SpiralIn.StartingStation] = ent.SpiralIn
            entities[ent.Arc.StartingStation] = ent.Arc
            entities[ent.SpiralOut.StartingStation] = ent.SpiralOut
        elif ent.Type == TLB.aeccSpiralTangentSpiralGroup:
            entities[ent.SpiralIn.StartingStation] = ent.SpiralIn
            entities[ent.Tangent.StartingStation] = ent.Tangent
            entities[ent.SpiralOut.StartingStation] = ent.SpiralOut
        elif ent.Type == TLB.aeccSpiralTangentGroup:
            entities[ent.SpiralIn.StartingStation] = ent.SpiralIn
            entities[ent.TangentOut.StartingStation] = ent.TangentOut
        elif ent.Type == TLB.aeccTangentSpiralGroup:
            entities[ent.TangentIn.StartingStation] = ent.TangentIn
            entities[ent.SpiralOut.StartingStation] = ent.SpiralOut
        elif ent.Type == TLB.aeccSpiralCurveGroup:
            entities[ent.SpiralIn.StartingStation] = ent.SpiralIn
            entities[ent.ArcOut.StartingStation] = ent.ArcOut
        elif ent.Type == TLB.aeccTCurveSpiralGroup:
            entities[ent.ArcIn.StartingStation] = ent.ArcIn
            entities[ent.SpiralOut.StartingStation] = ent.SpiralOut

    # Sort Alignment entities by station
    entities = OrderedDict(sorted(entities.items(), key=lambda t: t[0]))

    # Make sure each entity starts where the previous one ends
    values = entities.values()
    for i in range(len(values) - 1):
        assert isalmostequal(values[i].EndingStation,
                             values[i + 1].StartingStation)

    # Add applicable entity starting stations to pointStations
    for station in entities.keys():
        append = False
        if (station >= STARTING_STATION) and \
           (station <= ENDING_STATION) and not \
           isnuminiterable(station, pointStations):
                append = True

        # Dealig with possible inconsistency that might be caused due to
        # alignment's segment being too small. Added an extra "if" statement
        # here instead of a comparison above so that we can take action in
        # case of such an alignment (at the moment a warning is issued).
        # TODO: Maybe we need to introduce a different control constant here
        # such as PI_MIN_DIST for example.
        if isnuminiterable(station, pointStations, TOO_CLOSE):
            msg = "WARNING: Station %.2f too close with " % (station) + \
                "adjucent station!\nThis is a sign that alignment needs " + \
                "refinement. Ommiting the above station..."
            issuewarning(msg)
            append = False

        if append:
            pointStations.append(station)

# Get desired alignment profile
numProfiles = len(alignment.Profiles)
if numProfiles == 1:
    profile = alignment.Profiles[0]
elif numProfiles > 1:
    profiles = dict([(i.Name, i) for i in alignment.Profiles])
    while True:
        try:
            profile = profiles[doc.Utility.GetString(
                False, "Select profile (%s):" % (" or ".join(profiles.keys())))]
            break
        except KeyError:
            continue
else:
    issuewarning("WARNING: Alignment has no profile data!\nExiting...")
    profile = None

# Get alignment profile PVI stations if needed
if POINTS_AT_PVI_STATIONS:
    assert profile is not None
    for pvi in profile.PVIs:
        station = pvi.Station
        if (station >= STARTING_STATION) and \
           (station <= ENDING_STATION) and not \
           isnuminiterable(station, pointStations, TOO_CLOSE):  # Do we need to take action in case
                pointStations.append(station)                   # PVI is close with alignment's PI?

# Append POINT_MANDATORY_STATIONS in pointStations. After that it
# will be possible that 2 stations in pointStations will be too close
# with each other.
for station in POINT_MANDATORY_STATIONS:
    if (station >= STARTING_STATION) and \
       (station <= ENDING_STATION) and not \
       isnuminiterable(station, pointStations):
            pointStations.append(station)

pointStations.sort()

# So far we have a list of stations where points should be created.
# Now we should begin to interpolate between them according to STEP
i = 0
while i != len(pointStations) - 1:
    prevStation = pointStations[i]
    nextStation = pointStations[i + 1]
    if nextStation - prevStation > STEP + TOLERANCE:
        pointStations.append(prevStation + STEP)
        pointStations.sort()
    i += 1

# Check if there are any stations in pointStations that are too close
# due to POINT_MANDATORY_STATIONS
for i in range(len(pointStations) - 1):
    if isalmostequal(pointStations[i], pointStations[i + 1], TOO_CLOSE):
        issuewarning("WARNING: Stations %f and %f are too close!" % (
            pointStations[i], pointStations[i + 1]))

# Draw 3D Polylines at givven stations and offsets
for offset in OFFSETS:
    command = ["3dpoly"]
    print 70 * "-"
    for station in pointStations:
        print "Point at station %.6f - offset %.2f" % (station, offset)
        x, y = alignment.PointLocation(station, offset)
        try:
            z = profile.ElevationAt(station)
        except COMError:  # raised when <station> is out of <profile> range
            z = 0.0  # TODO: Maybe use a constant here (i.e. DEFAULT_ELEV)?
        except AttributeError:  # raised when alignment has no profile
            z = 0.0
        command.append("%s,%s,%s" % (x, y, z))

    command.append(" ")
    doc.SendCommand(" ".join(command))
