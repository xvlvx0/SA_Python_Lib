"""
This example creates a Plane Relationship as a base task.
Additionally it can create a callout.
"""
import sys
sys.path.append('C:\Analyzer Data\Scripts\SAPython\lib')
import SAPyLib as sa


class makeRelationshipPlane():
    """A class for creating a plane relationship."""

    _created_ = False       # class variable if the make() method was succesfull
    RelCollection = None    # the active collection
    RelName = None          # the relationship name [str]
    MeasuredCollection = None  # the collection name for the measured pointa
    MeasuredGroup = None    # the group name for the measured points
    NomCollection = None    # Nominals collection name
    NomObj = None           # the nominal object
    CallOutName = None      # the callout name

    def __init__(self, *args, **kwargs):
        """Initialize the class."""
        if 'relcollection' in kwargs:
            self.RelCollection = kwargs['relcollection']

        if 'relname' in kwargs:
            self.RelName = kwargs['relname']

        if 'measuredcollection' in kwargs:
            self.MeasuredCollection = kwargs['measuredcollection']

        if 'measuredgroup' in kwargs:
            self.MeasuredGroup = kwargs['measuredgroup']

        if 'nomcollection' in kwargs:
            self.NomCollection = kwargs['nomcollection']

        if 'nomobj' in kwargs:
            self.NomObj = kwargs['nomobj']

    def make(self):
        """Make the relationship."""
        
        # set default collection
        sa.set_or_construct_default_collection(
            self.RelCollection
        )

        # work around for adding measured group to relationship
        sa.construct_a_point('', self.MeasuredGroup, 'temp', 0.0, 0.0, 0.0)

        # make the relationship
        sa.make_geometry_fit_and_compare_to_nominal_relationship(
            '', self.RelName,
            self.NomCollection, self.NomObj,
            [("", self.MeasuredGroup), ],
            '', ''
        )
        self._created_ = True

        sa.set_relationship_desired_meas_count('', self.RelName, 0)
        sa.set_geom_relationship_criteria('', self.RelName, 'Flatness')
        # sa.set_geom_relationship_criteria('', self.RelName, 'Centroid Z')
        # sa.set_geom_relationship_criteria('', self.RelName, 'Avg Dist Between')

    def addCallout(self):
        """Add a callout view for the relationship."""
        if not self._created_:
            raise IOError("Relationship isn't created yet.")

        self.CallOutName = f'Callout_{self.RelName}'
        # Create Relationship Callout
        sa.create_relationship_callout(
            self.CallOutName,
            )

        # Set Default Callout View Properties
        sa.set_default_callout_view_properties(
            calloutname=self.CallOutName,
            )

        # Create Text Callout
        sa.create_text_callout('Info on callout')


if __name__ == '__main__':
    collection = 'MyCollection'         # distination collection 
    relname = 'MyPlane'                 # relationshipname
    measuredgroup = 'MyMeasuredPoints'  # The group containing the measured points

    # set default collection
    sa.set_or_construct_default_collection('A')
    # delete collections
    sa.delete_collection(collection)

    # Make a plane relationship.
    # Inputs are:
    # - The destination collection name for the relationship
    # - The relationship name
    # - The measured point group name (can be an empty group)
    # - The nominal geometry [optional]
    myPlane = makeRelationshipPlane(
                relcollection=collection,
                relname=relname,
                measuredcollection=collection,
                measuredgroup=measuredgroup,
                nomcollection='Nominals',
                nomobj='NomPlane',
                )
    myPlane.make()
    # myPlane.addCallout()

    sys.exit(0)
