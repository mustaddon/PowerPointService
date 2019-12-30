using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace RandomSolutions
{
    public class PowerPointService
    {
        protected readonly Lazy<Dictionary<string, IPipeTransform>> Pipes;

        public PowerPointService(IEnumerable<IPipeTransform> pipes = null)
        {
            Pipes = new Lazy<Dictionary<string, IPipeTransform>>(() => (pipes ?? _defaultPipes())
                .GroupBy(x => x?.Name?.Trim().ToLower())
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .ToDictionary(g => g.Key, g => g.First()));
        }

        public virtual byte[] CreateFromTemplate(byte[] templatePresentation, Func<int, int, object> slideModelFactory)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(templatePresentation, 0, templatePresentation.Length);

                using (var doc = PresentationDocument.Open(ms, true))
                    _fillSlides(doc, slideModelFactory);

                return ms.ToArray();
            }
        }

        public virtual byte[] InsertSlides(byte[] sourcePresentation, byte[] targetPresentation, int targetInsertIndex = -1, Func<int, int, bool> sourceSlideSelector = null)
        {
            using (var targetStream = new MemoryStream())
            using (var sourceStream = new MemoryStream())
            {
                targetStream.Write(targetPresentation, 0, targetPresentation.Length);
                sourceStream.Write(sourcePresentation, 0, sourcePresentation.Length);

                using (var target = PresentationDocument.Open(targetStream, true))
                using (var source = PresentationDocument.Open(sourceStream, true))
                    _insertSlides(source, target, targetInsertIndex, sourceSlideSelector);

                return targetStream.ToArray();
            }
        }

        public virtual byte[] DeleteSlides(byte[] presentation, Func<int, int, bool> slideSelector)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(presentation, 0, presentation.Length);

                using (var doc = PresentationDocument.Open(ms, true))
                {
                    var slist = doc.PresentationPart.Presentation.SlideIdList;
                    var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();

                    for (var i = 0; i < slideIds.Length; i++)
                    {
                        if (!slideSelector(i, slideIds.Length))
                            continue;

                        var slideId = slideIds[i];
                        var slide = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);

                        slist.RemoveChild(slideId);
                        _cleanCustomShow(doc.PresentationPart.Presentation.CustomShowList, slideId.RelationshipId);
                        doc.PresentationPart.Presentation.Save();
                        doc.PresentationPart.DeletePart(slide);
                    }
                }

                return ms.ToArray();
            }
        }


        void _insertSlides(PresentationDocument source, PresentationDocument target, int targetInsertIndex, Func<int, int, bool> sourceSlideSelector)
        {
            if (target.PresentationPart.Presentation.SlideIdList == null)
                target.PresentationPart.Presentation.SlideIdList = new SlideIdList();

            var slideMasterPartsMap = _cloneSlideMasterParts(source, target);
            var targetSlidesCount = target.PresentationPart.Presentation.SlideIdList.Count();
            var index = targetInsertIndex < 0 ? Math.Max(0, targetSlidesCount + targetInsertIndex + 1) : Math.Min(targetSlidesCount, targetInsertIndex);
            var nextId = _getMaxSlideId(target.PresentationPart.Presentation.SlideIdList) + 1;
            var sourceSlideIds = source.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToArray();

            for (var i = 0; i < sourceSlideIds.Length; i++)
            {
                if (sourceSlideSelector?.Invoke(i, sourceSlideIds.Length) == false)
                    continue;

                var sourceSlideId = sourceSlideIds[i];
                var sourceSlidePart = (SlidePart)source.PresentationPart.GetPartById(sourceSlideId.RelationshipId);
                var sourceSlideLayoutPartId = sourceSlidePart.SlideLayoutPart.SlideMasterPart.GetIdOfPart(sourceSlidePart.SlideLayoutPart);
                var targetSlideMasterPart = slideMasterPartsMap[sourceSlidePart.SlideLayoutPart.SlideMasterPart];
                var targetSlideLayoutPart = (SlideLayoutPart)targetSlideMasterPart.GetPartById(sourceSlideLayoutPartId);//.SlideLayoutParts.First(x => x.SlideLayoutParts).FirstOrDefault(x => x.SlideLayout.CommonSlideData.Name.Value.IndexOf(sourceSlidePart.SlideLayoutPart.SlideLayout.CommonSlideData.Name, StringComparison.InvariantCultureIgnoreCase) >= 0)
                var targetSlidePart = target.PresentationPart.AddPart(sourceSlidePart);

                targetSlidePart.DeleteParts(targetSlidePart.Parts.Select(x => x.OpenXmlPart)
                    .Where(x => x == targetSlidePart.SlideLayoutPart
                        || x.GetType() == typeof(NotesSlidePart)));

                targetSlidePart.AddPart(targetSlideLayoutPart);
                var targetSlideId = target.PresentationPart.Presentation.SlideIdList.InsertAt(
                    new SlideId() { Id = nextId++, RelationshipId = target.PresentationPart.GetIdOfPart(targetSlidePart) },
                    index++);
            }
        }

        void _fillSlides(PresentationDocument doc, Func<int, int, object> modelFactory)
        {
            var slist = doc.PresentationPart.Presentation.SlideIdList;
            var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();
            var nextId = _getMaxSlideId(slist) + 1;

            for (var i = 0; i < slideIds.Length; i++)
            {
                var sm = modelFactory(i, slideIds.Length);

                if (sm == null)
                    continue;

                var slideId = slideIds[i];
                var slide = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
                var models = sm as IEnumerable<object> ?? new object[] { sm };
                var prevSlideId = slideId;

                foreach (var model in models)
                {
                    var newSlidePart = doc.PresentationPart.AddNewPart<SlidePart>();
                    newSlidePart.Slide = new Slide(_insertValues(model, slide.Slide.OuterXml));
                    _copyPartsAndRelationships(slide, newSlidePart);
                    prevSlideId = slist.InsertAfter(new SlideId() { Id = nextId++, RelationshipId = doc.PresentationPart.GetIdOfPart(newSlidePart) }, prevSlideId);
                }

                slist.RemoveChild(slideId);
                _cleanCustomShow(doc.PresentationPart.Presentation.CustomShowList, slideId.RelationshipId);
                doc.PresentationPart.Presentation.Save();
                doc.PresentationPart.DeletePart(slide);
            }
        }

        string _insertValues(object model, string xml)
        {
            if (model == null)
                return null;

            var result = _insertRows(model, xml);

            return Regex.Replace(result, _cmdPattern,
                x => _getValue(model, x.Groups[1].Value)?.ToString(),
                RegexOptions.IgnorePatternWhitespace);
        }

        string _insertRows(object model, string xml)
        {
            return Regex.Replace(xml, @"<a:tr.+?</a:tr>", x =>
            {
                var itemsPath = Regex.Matches(x.Value, _cmdPattern, RegexOptions.IgnorePatternWhitespace).Cast<Match>()
                    .Select(xx => _getRowSourcePath(model, xx.Groups[1].Value))
                    .FirstOrDefault(xx => xx != null);

                if (itemsPath == null)
                    return x.Value;

                var items = _getObjValue(model, itemsPath) as System.Collections.IEnumerable;
                var result = new StringBuilder();

                if (items != null)
                    foreach (var item in items)
                        result.Append(Regex.Replace(x.Value, _cmdPattern,
                            xx =>
                            {
                                var cmd = _removeTags(xx.Groups[1].Value).Trim();
                                return cmd.StartsWith(itemsPath)
                                    ? _getValue(item, cmd.Substring(Math.Min(itemsPath.Length + 1, cmd.Length)))?.ToString()
                                    : xx.Value;
                            },
                            RegexOptions.IgnorePatternWhitespace));

                return result.ToString();
            });
        }

        string _getRowSourcePath(object obj, string cmd)
        {
            var propNames = _removeTags(cmd).Split('|').First().Trim().Split('.');
            var type = obj.GetType();

            for (var i = 0; i < propNames.Length - 1; i++)
            {
                var pi = type.GetProperty(propNames[i]);

                if (pi == null)
                    break;

                if (_isArray(pi.PropertyType))
                    return string.Join(".", propNames.Take(i + 1));

                if (obj == null)
                    break;

                type = pi.PropertyType;
                obj = pi.GetValue(obj);
            }

            return null;
        }

        object _getValue(object obj, string cmd)
        {
            var parts = _removeTags(cmd).Split('|');
            var value = _getObjValue(obj, parts.First());

            for (var i = 1; i < parts.Length; i++)
            {
                var pipe = parts[i].Trim().Split(':').First().Trim().ToLower();

                if (Pipes.Value.ContainsKey(pipe))
                {
                    var args = Regex.Matches(parts[i], @"'([^']*?)'").Cast<Match>().Select(x => x.Groups[1].Value).ToArray();
                    value = Pipes.Value[pipe].Transform(value, args);
                }
            }

            return value;
        }

        static string _removeTags(string str)
        {
            return Regex.Replace(str, @"<[^<>]*?>", "");
        }

        static object _getObjValue(object obj, string path)
        {
            try
            {
                var param = Expression.Parameter(obj.GetType(), string.Empty);
                var prop = path.Trim().Split('.').Aggregate<string, Expression>(param, (r, x) => Expression.PropertyOrField(r, x));
                var getter = Expression.Lambda(prop, param);
                return getter.Compile().DynamicInvoke(obj);
            }
            catch { }

            return null;
        }

        static void _copyPartsAndRelationships(SlidePart source, SlidePart target)
        {
            source.Parts?.Where(x => x.OpenXmlPart.GetType() != typeof(NotesSlidePart)).ToList()
                .ForEach(x => target.AddPart(x.OpenXmlPart, x.RelationshipId));

            source.HyperlinkRelationships?.ToList()
                .ForEach(x => target.AddHyperlinkRelationship(x.Uri, x.IsExternal, x.Id));

            source.ExternalRelationships?.ToList()
                .ForEach(x => target.AddExternalRelationship(x.RelationshipType, x.Uri, x.Id));

            source.DataPartReferenceRelationships?.ToList()
                .ForEach(x =>
                {
                    if (x is AudioReferenceRelationship)
                        target.AddAudioReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
                    else if (x is VideoReferenceRelationship)
                        target.AddVideoReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
                    else if (x is MediaReferenceRelationship)
                        target.AddMediaReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
                });
        }

        static void _cleanCustomShow(CustomShowList customShowList, string slideRelId)
        {
            if (customShowList != null)
                foreach (var customShow in customShowList.Elements<CustomShow>())
                    if (customShow.SlideList != null)
                    {
                        var slideListEntries = new LinkedList<SlideListEntry>();

                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                slideListEntries.AddLast(slideListEntry);

                        foreach (var slideListEntry in slideListEntries)
                            customShow.SlideList.RemoveChild(slideListEntry);
                    }
        }

        static Dictionary<SlideMasterPart, SlideMasterPart> _cloneSlideMasterParts(PresentationDocument source, PresentationDocument target)
        {
            var nextId = _getMaxSlideMasterId(target.PresentationPart.Presentation.SlideMasterIdList) + 1;
            var mapping = new Dictionary<SlideMasterPart, SlideMasterPart>();

            foreach (var sourceSlideMasterId in source.PresentationPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>())
            {
                var sourceSlideMasterPart = (SlideMasterPart)source.PresentationPart.GetPartById(sourceSlideMasterId.RelationshipId);
                var targetSlideMasterPart = target.PresentationPart.AddPart(sourceSlideMasterPart);
                var targetSlideMasterId = new SlideMasterId() { Id = nextId++, RelationshipId = target.PresentationPart.GetIdOfPart(targetSlideMasterPart) };
                target.PresentationPart.Presentation.SlideMasterIdList.Append(targetSlideMasterId);
                mapping.Add(sourceSlideMasterPart, targetSlideMasterPart);
            }

            foreach (var slideMasterPart in target.PresentationPart.SlideMasterParts)
            {
                if (slideMasterPart.SlideMaster.SlideLayoutIdList != null)
                    foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                        slideLayoutId.Id = nextId++;

                slideMasterPart.SlideMaster.Save();
            }

            return mapping;
        }

        static uint _getMaxSlideId(SlideIdList slideIdList)
        {
            // Slide identifiers have a minimum value of greater than or
            // equal to 256 and a maximum value of less than 2147483648. 
            return Math.Max(256, slideIdList?.Elements<SlideId>().Max(x => x.Id) ?? 0);
        }

        static uint _getMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
        {
            // Slide master identifiers have a minimum value of greater than
            // or equal to 2147483648. 
            return Math.Max(2147483648, slideMasterIdList?.Elements<SlideMasterId>().Max(x => x.Id) ?? 0);
        }

        static bool _isArray(Type type)
        {
            return type != null && type != typeof(string) && typeof(System.Collections.IEnumerable).IsAssignableFrom(type);
        }

        static IEnumerable<IPipeTransform> _defaultPipes()
        {
            try
            {
                return typeof(IPipeTransform).Assembly.GetTypes()
                    .Where(x => !x.IsAbstract
                        && typeof(IPipeTransform).IsAssignableFrom(x)
                        && x.GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, new Type[0], null) != null)
                    .Select(x => Activator.CreateInstance(x) as IPipeTransform);
            }
            catch { }

            return new IPipeTransform[0];
        }

        static string _cmdPattern = @"{{(.*?)}}";

    }
}
