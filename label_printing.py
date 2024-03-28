#region Imports
import csv
import json
from pathlib import Path
from subprocess import run
from tempfile import NamedTemporaryFile
from tkinter import messagebox
from typing import Self

from natsort import natsort_keygen, natsorted
from qrcode import make
from qrcode.main import GenericImage
#endregion

#region Constants
# Make sure keys match structure order: 'https://learn.microsoft.com/en-us/windows/win32/printdocs/printer-info-2#Syntax'
PRINTER_INFO_2_ENUM = ['pServerName', 'pPrinterName', 'pShareName', 'pPortName', 'pDriverName', 'pComment',
					   'pLocation', 'pDevMode', 'pSepFile', 'pPrintProcessor', 'pDatatype', 'pParameters',
					   'pSecurityDescriptor', 'Attributes', 'Priority', 'DefaultPriority', 'StartTime',
					   'UntilTime', 'Status', 'cJobs', 'AveragePPM']
CJI = PRINTER_INFO_2_ENUM.index('cJobs')
"""Current job index."""
TAG_DELIMITER = '|'
"""Delimiter character for multiple skipped/forced tags."""
#endregion


class Tag:
	""" Represents a single sample tag within a geological sampling context.
		A tag denotes a specific sample taken from a hole, marked by its unique name and its depth range (starting to ending depths).

		### Properties
		- hole: `str` -- The identifier for the hole from which the sample was taken.
		- name: `str` -- The name or identifier of the tag.
		- starting_depth: `float` -- The starting depth of the sample, in meters.
		- ending_depth: `float` -- The ending depth of the sample, in meters.
	"""

	def __init__(self, hole: str, name: str, **kwargs: float | str) -> None:
		""" Initializes a new instance of the `Tag` class, representing a single sample tag.

			### Parameters
			- hole: `str` -- The identifier for the hole from which the sample was taken.
			- name: `str` -- The name or identifier of the tag.

			### Keyword Arguments
			- starting_depth: `float | str` -- The starting depth of the sample, in meters.
			- ending_depth: `float | str` -- The ending depth of the sample, in meters.
		"""

		self.__hole__ = hole
		self.__name__ = name
		self.__starting_depth__ = float(kwargs['starting_depth'])
		self.__ending_depth__ = float(kwargs['ending_depth'])

	def __repr__(self) -> str:
		return f'{self.hole}-{self.name} | {self.starting_depth:.2f} m - {self.ending_depth:.2f} m'

	def __str__(self) -> str:
		return f'{self.hole}-{self.name}'

	def __lt__(self, other: Self):
		return self.starting_depth < other.starting_depth

	@property
	def hole(self) -> str:
		return self.__hole__

	@property
	def name(self) -> str:
		return self.__name__

	@property
	def starting_depth(self) -> float:
		return self.__starting_depth__

	@property
	def ending_depth(self) -> float:
		return self.__ending_depth__

class Box:
	""" Represents a box that contains multiple `Tags`.
		Each box is associated with a particular hole and can contain several tags, representing samples.
		The `Box` class manages inclusion of tags based on various criteria such as depth and whether specific tags are forced to be included or skipped.

		### Properties
		- hole: `str` -- The identifier for the hole associated with this box.
		- name: `str` -- The name or identifier of the box.
		- starting_depth: `float` -- The starting depth range for samples within this box, in meters.
		- ending_depth: `float` -- The ending depth range for samples within this box, in meters.
		- tags: `list[Tag]` -- A list of `Tag` instances that belong to this box.
		- skipped_tags: `set[str]` -- A set of tag names that should not be included in the box.
		- forced_tags: `set[str]` -- A set of tag names that must be included in the box regardless of other criteria.
		- tag_at_sample_start: `bool` -- Indicates whether the tag is located at the sample's starting or ending depth.
	"""

	def add_tag(self, tag: Tag) -> None:
		""" Attempts to add a `Tag` to this `Box` if it meets certain criteria related to hole matching, skipped and forced tags, and depth ranges.

			### Parameters
			- tag: `Tag` -- The `Tag` to potentially add to the Box.

			### Note
			May invoke a GUI dialog for edge cases.
		"""

		if (tag.hole != self.hole) or (tag.name in self.skipped_tags):
			return

		if tag.name not in self.forced_tags:
			depth = tag.starting_depth if self.tag_at_sample_start else tag.ending_depth
			if not (self.starting_depth <= depth <= self.ending_depth):
				return

			if (depth == self.starting_depth) or (depth == self.ending_depth):
				message = f'Include tag "{tag}" in box "{self}"?'
				include_edge_case = messagebox.askyesno(title='Edge Case', message=message)
				if not include_edge_case:
					return

		self.tags.append(tag)

	def __init__(self, hole: str, name: str, **kwargs: float | str) -> None:
		""" Initializes a new instance of the `Box` class, representing a container for multiple `Tags`.

			### Parameters
			- hole: `str` -- The identifier for the hole from which the samples in the box were taken.
			- name: `str` -- The name or identifier of the box.

			### Keyword Arguments
			- starting_depth: `float | str` -- The starting depth of the box's sample range, in meters.
			- ending_depth: `float | str` -- The ending depth of the box's sample range, in meters.
		"""

		self.__hole__ = hole
		self.__name__ = name
		self.__starting_depth__ = float(kwargs['starting_depth'])
		self.__ending_depth__ = float(kwargs['ending_depth'])

		self.__tags__ = []
		self.__skipped_tags__ = {}
		self.__forced_tags__ = {}
		self.__tag_at_sample_start__ = True

	def __repr__(self) -> str:
		return f'{self.hole}-{self.name} | {self.starting_depth:.2f} m - {self.ending_depth:.2f} m'

	def __str__(self) -> str:
		return f'{self.hole}-{self.name}'

	def __lt__(self, other: Self) -> bool:
		return (self.starting_depth < other.starting_depth) if self.hole == other.hole else (self.hole == natsorted((self.hole, other.hole))[0])

	@property
	def hole(self) -> str:
		return self.__hole__

	@property
	def name(self) -> str:
		return self.__name__

	@property
	def starting_depth(self) -> float:
		return self.__starting_depth__

	@property
	def ending_depth(self) -> float:
		return self.__ending_depth__

	@property
	def tags(self) -> list[Tag]:
		return self.__tags__

	@property
	def skipped_tags(self) -> set[str]:
		return self.__skipped_tags__
	@skipped_tags.setter
	def skipped_tags(self, value: set[str]) -> None:
		self.__skipped_tags__ = value

	@property
	def forced_tags(self) -> set[str]:
		return self.__forced_tags__
	@forced_tags.setter
	def forced_tags(self, value: set[str]) -> None:
		self.__forced_tags__ = value

	@property
	def tag_at_sample_start(self) -> bool:
		return self.__tag_at_sample_start__
	@tag_at_sample_start.setter
	def tag_at_sample_start(self, value: bool) -> None:
		self.__tag_at_sample_start__ = bool(value)

class Printer:
	""" Manages the printing process for labels, configuring print commands based on the specified label size.
		This class encapsulates the details required to print to a specific printer, handling the construction of command-line print commands.

		### Properties
		- name: `str` -- The name of the printer. Currently hardcoded to 'DYMO LabelWriter 450'.
		- command: `str` -- The command used to print to the printer. It includes parameters tailored to the label size and printer settings.
	"""

	def print(self, filepath: Path | str) -> None:
		""" Sends a print command to the printer for the file located at the specified filepath.

			### Parameters
			- filepath: `Path | str` -- The path to the file to be printed.
		"""

		if not isinstance(filepath, Path):
			filepath = Path(filepath)

		run(f'{self.command} "{filepath}"')

	def __init__(self, label_size: str) -> None:
		""" Initializes a new instance of the Printer class, configuring it based on the specified label size.
			This setup includes determining the appropriate paper size and constructing the command used for printing

			### Parameters
			- label_size: `str` -- The size of the label to be printed. Supported values are 'small' and 'large'.

			### Raises
			- `ValueError` if an unsupported label size is specified.\
				This exception indicates that the label size provided does not match any of the predefined configurations ('small' or 'large') and thus cannot be processed.
		"""

		executable = Path(r'bin\SumatraPDF-3.5.2-64.exe').absolute()
		match label_size.casefold():
			case 'small':
				paper = '30252 Address'
			case 'large':
				paper = '30323 Shipping'
			case _:
				raise ValueError('Unsupported paper type!')

		self.__name__ = 'DYMO LabelWriter 450'
		self.__command__ = f'"{executable}" -print-to "{self.name}" -print-settings "noscale,paper={paper}" -silent -exit-when-done'

	@property
	def name(self) -> str:
		return self.__name__

	@property
	def command(self) -> str:
		return self.__command__


def generate_latex(imagepath: Path | str, qrcode_data: str) -> Path:
	""" Generates a LaTeX document from a given image path and QR code data, tailored to a specified label size.

		### Parameters
		- imagepath: `Path | str` -- The path to the QR code image file to be included in the LaTeX document.
		- qrcode_data: `str` -- Data contained within the QR code.

		### Returns
		- `Path` -- Generated TEX filepath.
	"""

	imagepath = imagepath if isinstance(imagepath, Path) else Path(imagepath)
	with NamedTemporaryFile('wt', newline='', suffix='.tex', delete=False) as texfile:
		texfile.write('\\documentclass{article}\n')
		texfile.write('\\usepackage[export]{adjustbox}\n')
		texfile.write('\\usepackage{float}\n')

		match label_size:
			case 'small':
				texfile.write('\\usepackage[margin=0mm, left=1mm, right=1mm, top=4mm, bottom=4mm, paperwidth=28mm, paperheight=89mm]{geometry}\n')
			case 'large':
				texfile.write('\\usepackage[margin=0mm, left=4mm, right=1mm, top=1mm, bottom=1mm, paperwidth=59mm, paperheight=102mm]{geometry}\n')

		texfile.write('\\usepackage{graphicx}\n')
		texfile.write('\\usepackage{pdflscape}\n')
		texfile.write('\\usepackage[scaled]{beramono}\n')
		texfile.write('\\renewcommand*\\familydefault{\\ttdefault}\n')
		texfile.write('\\usepackage[T1]{fontenc}\n')
		texfile.write('\\begin{document}\n')
		texfile.write('\\begin{landscape}\n')
		texfile.write('\\noindent\n')

		match label_size:
			case 'small':
				texfile.write('\\begin{minipage}{30mm}\n')
				texfile.write(f'\\includegraphics[width=25mm, height=25mm]{{{imagepath.name}}}\n')
				texfile.write('\\end{minipage}\n')
				texfile.write('\\hspace{-7.5mm}\n')
				texfile.write('\\begin{minipage}{50mm}\n')
			case 'large':
				texfile.write('\\begin{minipage}{55mm}\n')
				texfile.write(f'\\includegraphics[width=50mm, height=50mm]{{{imagepath.name}}}\n')
				texfile.write('\\end{minipage}\n')
				texfile.write('\\hspace{-7.5mm}\n')
				texfile.write('\\begin{minipage}{45mm}\n')

		texfile.write('\\begin{adjustbox}{max width=\\textwidth}\n')
		texfile.write('\\centering\n')
		texfile.write('\\begin{tabular}{c c}\n')

		markers = qrcode_data.splitlines()
		header = markers.pop(0).split(',')
		header[-1] = ('Samples starts at tags' if header[-1].casefold() == 'true' else 'Samples ends at tags') if tags_enabled else ''

		texfile.write(f"\\multicolumn{{2}}{{c}}{{\\large\\textbf{{{','.join(header[:-1])}}}\\par}} \\\\\n")
		texfile.write(f"\\multicolumn{{2}}{{c}}{{\\large\\textbf{{{header[-1]}}}\\par}} \\\\\n")

		match label_size:
			case 'small':
				if len(markers) > 10:
					temp = markers[:4]
					temp.append('\\cdots')
					temp.append('\\cdots')
					temp += markers[-4:]
					markers = temp
			case 'large':
				if len(markers) > 26:
					temp = markers[:12]
					temp.append('\\cdots')
					temp.append('\\cdots')
					temp += markers[-12:]
					markers = temp

		for i, sample in enumerate(markers):
			if not i%2:
				texfile.write(sample if (i + 1) != len(markers) else f'{sample} &\n')
			else:
				texfile.write(f' & {sample} \\\\\n' if (i + 1) != len(markers) else f' & {sample}\n')

		texfile.write('\\end{tabular}\n')
		texfile.write('\\end{adjustbox}\n')
		texfile.write('\\end{minipage}\n')
		texfile.write('\\end{landscape}\n')
		texfile.write('\\end{document}\n')

		return Path(texfile.name)


if __name__ == '__main__':
	# small (30252): 28*89 mm² | large (30323): 59*102 mm²
	label_size = 'small'
	printer = Printer(label_size)

	labels_datapath = Path(r'examples\Labels.csv')
	tags_datapath = Path(r'examples\Samples.csv')
	printed_datapath = labels_datapath.with_name('Printed Labels.json')

	label_keys = {'hole': 'Hole ID',
				  'box': 'Box ID',
				  'box_start': 'From',
				  'box_stop': 'To',
				  'tag_position': 'Tag Position',
				  'skipped_tags': 'Skipped Tags',
				  'forced_tags': 'Forced Tags'}

	tags_enabled = True
	tag_keys = {'hole': 'Hole',
				'tag': 'Sample',
				'tag_start': 'From',
				'tag_stop': 'To'}

	# Read data about already printed labels
	printed_labels: dict[str, list[str]] = {}
	if printed_datapath.exists() and printed_datapath.stat().st_size:
		with open(printed_datapath, 'r') as jsonfile:
			printed_labels = json.load(jsonfile)

	# Read data about printable labels
	boxes: list[Box] = []
	with open(labels_datapath, 'r') as csvfile:
		reader = csv.DictReader(csvfile)
		for r in reader:
			if not r[label_keys['hole']]:
				continue

			box = Box(hole=r[label_keys['hole']],
					  name=r[label_keys['box']],
					  starting_depth=float(r[label_keys['box_start']]),
					  ending_depth=float(r[label_keys['box_stop']]))

			if (box.hole in printed_labels) and (box.name in printed_labels[box.hole]):
				continue

			if tags_enabled:
				box.tag_at_sample_start = r[label_keys['tag_position']] == tag_keys['tag_start']

				if ((st := label_keys['skipped_tags']) is not None) and any(st_list := r[st].split(TAG_DELIMITER)):
					box.skipped_tags = set(st_list)

				if ((ft := label_keys['forced_tags']) is not None) and any(ft_list := r[ft].split(TAG_DELIMITER)):
					box.forced_tags = set(ft_list)

			boxes.append(box)
	boxes.sort()

	# Read tag data
	if tags_enabled:
		labelled_holes = {b.hole for b in boxes}
		with open(tags_datapath, 'r') as csvfile:
			reader = csv.DictReader(csvfile)
			for r in reader:
				if (not r[tag_keys['hole']]) or (r[tag_keys['hole']] not in labelled_holes):
					continue

				tag = Tag(hole=r[tag_keys['hole']],
						  name=r[tag_keys['tag']],
						  starting_depth=float(r[ts]) if (ts := tag_keys['tag_start']) in r else 0.0,
						  ending_depth=float(r[ts]) if (ts := tag_keys['tag_stop']) in r else 0.0)
				[b.add_tag(tag) for b in boxes if b.hole == tag.hole]
		[b.tags.sort() for b in boxes]

	for box in boxes:
		qrcode_data = f"{box.hole},{box.name},{box.starting_depth:0.2f},{box.ending_depth:0.2f},{box.tag_at_sample_start}"
		for tag in box.tags:
			qrcode_data += f"\r\n{tag.name},{tag.starting_depth if box.tag_at_sample_start else tag.ending_depth:0.2f}"
		qrcode_image: GenericImage = make(qrcode_data)

		with NamedTemporaryFile('wb', suffix='.png', delete=False) as pngfile:
			qrcode_image.save(pngfile)
			texpath = generate_latex(pngfile.name, qrcode_data)

			run(f'lualatex --interaction=nonstopmode --enable-write18 "{texpath.stem}"', cwd=texpath.parent)
			printer.print(texpath.with_suffix('.pdf'))

		if box.hole in printed_labels:
			printed_labels[box.hole].append(box.name)
		else:
			printed_labels[box.hole] = [box.name]

	[printed_labels[hole].sort(key=natsort_keygen(lambda name: name)) for hole in printed_labels]
	with open(printed_datapath, 'w') as jsonfile:
		json.dump(printed_labels, jsonfile, indent='\t', sort_keys=True)
