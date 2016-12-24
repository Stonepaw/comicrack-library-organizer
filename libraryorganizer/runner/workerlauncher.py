
from ComicRack import ComicRack
from loworkerform import WorkerForm

worker = WorkerForm([],[], ComicRack())
worker.Shown -= worker.WorkerFormLoad
worker.ShowDialog()