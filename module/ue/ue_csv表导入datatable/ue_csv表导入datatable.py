import unreal

# Correctly load the blueprint class
blueprint_class = unreal.EditorAssetLibrary.load_blueprint_class(
    '/Game/LogicRes/FeatureTest/AudioTest/BP_AudioTest.BP_AudioTest')

# Check if the class was loaded successfully
if blueprint_class is None:
    unreal.log("Failed to load blueprint class.")
else:
    # Get the Class Default Object (CDO) of the blueprint
    cdo = blueprint_class.get_default_object()

    # unreal.log(str(dir(cdo)))

    cdo.call_method("printa")




