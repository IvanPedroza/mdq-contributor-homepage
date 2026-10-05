'use strict';

const links = {
  form: 'https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR__QMmLK2ntOv7m47E7ymNJUME1YTzhSVjZTNENGVjY1UlFVUU9ZWlhZNCQlQCN0PWcu',
  tracker: 'https://microsoft.sharepoint.com/teams/MDQContentPublication/Lists/Blog%20Publication%20Tracker/AllItems.aspx',
  guide: 'resources/publication-guide.docx',
  supplied: [
    {name:'Workbook 1 — title to confirm',url:'https://microsoft.sharepoint.com/:x:/t/MDQContentPublication/cQp7LBCu5B9MQ67hZQRTspnQEgUBvw4G8UXjDfwdeLMuy0EDIA'},
    {name:'Workbook 2 — title to confirm',url:'https://microsoft.sharepoint.com/:x:/t/MDQContentPublication/cQq2xyOVYh3eRLXReQGbKojGEgUBa0z5FDhoo457RAtGYJSO3Q'},
    {name:'Presentation — title to confirm',url:'https://microsoft.sharepoint.com/:p:/t/MDQContentPublication/cQr-lZR5Ut3VQ5QKhHz-YUubEgUBsuBzgtJ4PTr94BQm17q0kg'},
    {name:'Word document — title to confirm',url:'https://microsoft.sharepoint.com/:w:/t/MDQContentPublication/cQrBjL3xiV2ISonbAB2pMtPgEgUB1bSi0cfm8p6yqUbfJyDtqA'}
  ]
};
const $ = selector => document.querySelector(selector);
const escapeHTML = value => String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const externalLink = (url,label,classes='') => `<a class="${classes}" href="${url}" target="_blank" rel="noopener noreferrer">${label}</a>`;
$('#submit-link').href = links.form;
$('#workspace-link').href = links.tracker;
document.querySelectorAll('.tracker-link').forEach(a=>a.href=links.tracker);

// No demonstration records. Live records require an authorized SharePoint connection.
// In the native SharePoint build this panel is supplied by the List web part.
let articles = [];
let trackerConnected = false;
function renderArticles(){
  const query=$('#article-search').value.trim().toLowerCase(),stage=$('#stage-filter').value;
  const rows=articles.filter(item=>(!stage||item.stage===stage)&&`${item.title} ${item.author} ${item.stage} ${item.next}`.toLowerCase().includes(query));
  $('#article-rows').replaceChildren();
  $('#calendar-view').replaceChildren();
  $('#article-search').disabled=!trackerConnected;
  $('#stage-filter').disabled=!trackerConnected;
  $('#article-empty').hidden=trackerConnected&&rows.length>0;
  $('#article-empty').textContent=trackerConnected?'No articles match. Try another title or stage.':'Open the publication tracker for current article status and target dates.';
}
$('#article-search').addEventListener('input',renderArticles);
$('#stage-filter').addEventListener('change',renderArticles);
document.querySelectorAll('[data-view]').forEach(button=>button.addEventListener('click',()=>{
  document.querySelectorAll('[data-view]').forEach(b=>{const active=b===button;b.classList.toggle('selected',active);b.setAttribute('aria-pressed',String(active));});
  $('#list-view').hidden=button.dataset.view!=='list';
  $('#calendar-view').hidden=button.dataset.view!=='calendar';
}));
renderArticles();

const assetTypes={
  images:{title:'Readable in the published article.',specs:['PNG or JPG','Up to 1000 px wide','Alt text + caption'],text:'Supply original images with readable labels and approved captions. Use 1000 px, 650 px, or 205 px display widths as appropriate. Check sharpness again after saving the post.'},
  thumbnail:{title:'Make the first impression count.',specs:['800 × 400 px','PNG or JPG','Summary: 140 characters max'],text:'Every article needs a thumbnail and a separate listing summary. An optional header image is 1300 × 500 px; do not use a GIF for the header.'},
  video:{title:'Plan hosting before staging.',specs:['Final MP4','Captions or transcript','16:9 if suitable'],text:'Upload the final video to your workspace and ask Operations to coordinate hosting. For suitable landscape video, use 1920 × 1080 with a matching preview image. Allow time for hosting and a playback check.'},
  interactive:{title:'Make the experience work for everyone.',specs:['Hosting review','Static alternative','Public-access test'],text:'Bring interactive figures to Operations early to confirm the supported hosting route. Include a static image, caption, placement, and release approval. Do not assume an interactive file can be attached directly to the post.'},
  downloads:{title:'A useful file is an accessible file.',specs:['Public destination','Descriptive link text','Release approval'],text:'Provide the final document, a clear description, intended placement, and approval for public release. Ask Operations to arrange a public destination and test access without a work-account sign-in.'}
};
function setAsset(key){
  const item=assetTypes[key];
  document.querySelectorAll('[data-asset]').forEach(b=>{const active=b.dataset.asset===key;b.setAttribute('aria-selected',String(active));b.tabIndex=active?0:-1;});
  $('#asset-detail').setAttribute('aria-labelledby',`tab-${key}`);
  $('#asset-detail').innerHTML=`<h3>${item.title}</h3><div class="asset-specs">${item.specs.map(s=>`<span class="asset-spec">${s}</span>`).join('')}</div><p>${item.text}</p>`;
}
document.querySelectorAll('[data-asset]').forEach(b=>b.addEventListener('click',()=>setAsset(b.dataset.asset)));
setAsset('images');

const stages=[
  {name:'Blog content submission form',owner:'Author',provide:'Your proposal, intended audience, main author, and the contribution the story will make.',happens:'Submit the program Form to start the single publication path.'},
  {name:'Draft submission',owner:'Author · automated intake',provide:'The working Word draft and supporting materials for the proposed article.',happens:'The workflow creates the SharePoint folder and tracker entry, emails the submitter, and creates the Teams chat.'},
  {name:'1st content review',owner:'Katie Zoller',provide:'A shared draft with a clear narrative, proposed audience, and supporting evidence.',happens:'Katie checks story and calendar fit, obtains any missing draft, and assigns technical reviewers.'},
  {name:'SEO review',owner:'Ivan Pedroza · Anja Kurup',provide:'A descriptive title, short summary, SEO copy, and proposed Discovery taxonomy tags.',happens:'Operations reviews search readiness. The author revises the title and metadata as needed.'},
  {name:'Technical review',owner:'Assigned technical reviewers',provide:'Reproducible technical detail, supporting references, and responses to Word comments.',happens:'Subject-matter reviewers check accuracy. Resolve comments and signal readiness in Teams.'},
  {name:'Comms review',owner:'MDQ PR lead',provide:'Confirmed messaging, partner or customer mentions, and any needed external approvals.',happens:'The PR lead checks positioning and messaging. Flag third-party approvals early.'},
  {name:'Legal review',owner:'Discovery CELA · via Operations',provide:'Public-release approvals, third-party permissions, and the reviewed draft with supporting assets.',happens:'Editorial coordinates legal review by email. Legal reviewers are not in the per-blog Teams chat.'},
  {name:'Final content review',owner:'Katie Zoller',provide:'Resolved comments and complete wording, figures, captions, credits, and links.',happens:'Katie confirms the final content. The approved Word document becomes the source for staging.'},
  {name:'Amplification plan',owner:'Editorial · August Laguio · PR',provide:'Final tags, social visuals, promotion inputs, and readiness to share the published article.',happens:'The team agrees the promotion plan and coordinates amplification with the publication date.'},
  {name:'Staging in Tech Community editor',owner:'Main author · Operations',provide:'The approved Word copy, final assets, thumbnail, summaries, tags, and author profile access.',happens:'The main author creates the draft shell. Operations stages it; the author checks the saved preview against the approved Word copy.'},
  {name:'Publication',owner:'Katie Zoller',provide:'Final preview confirmation and readiness for the agreed editorial-calendar date.',happens:'Katie completes the final check and publishes. The amplification plan starts the same day.'}
];
$('#process-tabs').innerHTML=stages.map((s,i)=>`<button role="tab" id="stage-${i}" aria-controls="process-detail" aria-selected="${i===0}" tabindex="${i===0?'0':'-1'}" data-stage="${i}"><span class="number">${String(i+1).padStart(2,'0')}</span>${s.name}</button>`).join('');
function setStage(index){
  const stage=stages[index];
  document.querySelectorAll('[data-stage]').forEach(b=>{const active=Number(b.dataset.stage)===index;b.setAttribute('aria-selected',String(active));b.tabIndex=active?0:-1;});
  $('#process-detail').setAttribute('aria-labelledby',`stage-${index}`);
  $('#process-detail').innerHTML=`<div><small>WHO OWNS IT</small><strong>${stage.owner}</strong></div><div><small>WHAT YOU PROVIDE</small><p>${stage.provide}</p></div><div><small>WHAT HAPPENS NEXT</small><p>${stage.happens}</p></div>`;
}
document.querySelectorAll('[data-stage]').forEach(b=>b.addEventListener('click',()=>setStage(Number(b.dataset.stage))));
setStage(0);
document.querySelectorAll('[role="tablist"]').forEach(group=>group.addEventListener('keydown',event=>{
  if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;
  const tabs=[...group.querySelectorAll('[role="tab"]')],current=tabs.indexOf(document.activeElement);
  if(current<0)return;event.preventDefault();
  const next=event.key==='Home'?0:event.key==='End'?tabs.length-1:(current+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;
  tabs[next].click();tabs[next].focus();
}));

const faqs=[
  {q:'Images look sharp in the editor but blurry after saving.',tags:'photos resolution sharpness compression',a:'Check the saved preview at the intended display width. Compare it with the original file, keep labels readable, and avoid enlarging a small source image. Share the original, saved-page screenshot, and display width with Operations.',resource:'asset-guide'},
  {q:'Videos crop or change size during playback.',tags:'video aspect ratio player embed',a:'Check the final video, preview image, and player at the same aspect ratio. Test the saved post through playback, not just the editor preview. Send Operations the hosting link, browser details, and a screenshot of the crop.',resource:'hosting'},
  {q:'Interactive figures need hosting and static alternatives.',tags:'interactive html figure zip hosting accessibility',a:'Confirm the hosting route with Operations before staging. Supply the interactive package, a static alternative, caption, intended placement, and public-release approval. Hosting support must be confirmed for the specific asset.',resource:'hosting'},
  {q:'Attachments or links aren’t accessible to public readers.',tags:'files public external permissions access downloads',a:'A work-account link may work for you and still fail for readers. Test while signed out or in a private browsing window. Ask Operations for a public destination; avoid linking to the internal working draft.',resource:'asset-guide'},
  {q:'I can open the Word document but not the platform draft.',tags:'draft author editor profile permissions access',a:'SharePoint document access and Tech Community draft access are separate. Send Katie your Tech Community username and profile URL, plus the draft link and the access error.',resource:'profile'},
  {q:'Tags, SEO, summaries, and captions—what goes where?',tags:'tags taxonomy seo description thumbnail summary caption alt text metadata',a:'Tags categorize the article. The SEO description supports search. The thumbnail summary introduces the listing and is limited to 140 characters. Captions explain figures; alt text describes their content for accessibility.',resource:'taxonomy'}
];
$('#faq-list').innerHTML=faqs.map((faq,i)=>`<details class="faq-item" data-faq="${i}"><summary>${faq.q}</summary><p>${faq.a}</p><button class="text-link" data-open="${faq.resource}">Open related guidance</button></details>`).join('');
$('#faq-search').addEventListener('input',()=>{
  const query=$('#faq-search').value.trim().toLowerCase();let count=0;
  document.querySelectorAll('[data-faq]').forEach(el=>{const item=faqs[Number(el.dataset.faq)],matches=`${item.q} ${item.a} ${item.tags}`.toLowerCase().includes(query);el.hidden=!matches;el.open=!!query&&matches;if(matches)count++;});
  $('#faq-empty').hidden=count>0;
});

const helpRequest='Article title and tracker link:\nCurrent stage:\nWorkspace or asset link:\nWhat I expected / what happened:\nError message or screenshot:\nTarget publication date / help needed by:';
const suppliedResources = () => `<ul class="provided-links">${links.supplied.map(item=>`<li>${externalLink(item.url,item.name)}</li>`).join('')}</ul>`;
const resources={
  'process-map':()=>`<h2 id="dialog-title">Blog Publication Path</h2><p>The program’s supplied process map. Every submission follows this sequence.</p><img src="resources/publication-path.jpg" alt="Blog Publication Path: submission form, draft submission with automated workspace, tracker, email and Teams chat creation, first content review, SEO review, technical review, comms review, legal review, final content review, amplification plan, staging in Tech Community editor, then publication." style="display:block;width:100%;height:auto;margin-top:20px">${externalLink('resources/publication-path.jpg','Open full-size process map','button')}`,

  guide:()=>`<h2 id="dialog-title">Publication guide</h2><p>Microsoft Discovery Blog Publication Guide · version 1.0 · updated September 29, 2026.</p><h3>Start with these essentials</h3><ol><li>Work in your article’s SharePoint folder. Keep the Word draft and all supporting assets together.</li><li>Use the article’s Teams chat for handoffs and questions; resolve review comments in Word.</li><li>Complete content, technical, communications, and legal review before staging.</li><li>Confirm the main author and profile permissions before creating the Tech Community draft.</li><li>Check the saved preview against the approved Word copy and test public links.</li></ol><p>After final content review, only clear typos may change without reapproval. Changes to captions, summaries, SEO copy, or approved content need approval again. The agreed calendar date sets the publication target.</p><a class="button" href="${links.guide}">Open the full Word guide</a><p class="notice">This local preview uses the supplied guide from the project folder. The SharePoint resource link and document owner still need to be confirmed.</p>`,
  'asset-guide':()=>`<h2 id="dialog-title">Prepare the complete asset package</h2><p>Requirements below come from the publication guide, version 1.0. Keep final files in your article workspace.</p><dl><dt>Thumbnail</dt><dd>800 × 400 px, PNG or JPG. Supply a separate listing summary of no more than 140 characters.</dd><dt>Header (optional)</dt><dd>1300 × 500 px, PNG or JPG. No GIF header.</dd><dt>Article images</dt><dd>PNG or JPG, no wider than 1000 px. Display widths: large 1000 px, medium 650 px, small 205 px. At least one body image is required.</dd><dt>Animated GIF</dt><dd>Maximum 9.8 MB. Permitted in the body or thumbnail.</dd><dt>Video</dt><dd>Final MP4. For suitable landscape content: 1920 × 1080 and a matching 16:9 preview image. Include captions or a transcript.</dd><dt>Interactive figures</dt><dd>Ask Operations to confirm the hosting route. Include a static alternative and descriptive caption.</dd><dt>Downloads</dt><dd>Final file, descriptive label, placement, and a destination that public readers can access.</dd></dl><h3>Include with every request</h3><ul><li>Article title, tracker link, workspace link, and target publication date.</li><li>Final filenames, placement, captions, alt text, and credit or permission for third-party material.</li><li>Public-release approval and any accessibility alternatives.</li></ul><button class="button" data-open="hosting">Prepare a hosting request</button>`,
  hosting:()=>`<h2 id="dialog-title">Get media hosting help</h2><p>Start with Publication Operations—<strong>Ivan Pedroza or Anja Kurup</strong>—in your article’s Teams chat. Operations coordinates the Tech Community hosting request.</p><h3>Send one complete package</h3><ol><li>Upload the final assets to the article workspace.</li><li>Include titles, captions, alt text, placement, and the desired publication date.</li><li>For video, include the final MP4, preview image, and captions or transcript.</li><li>For an interactive figure, include its package and a static image. Confirm support before relying on an embed.</li><li>Include public-release approval and third-party permissions.</li></ol><p>The guide estimates about four days for video hosting; confirm timing with Operations for your request.</p><p class="notice">The guide gives conflicting descriptions of interactive hosting. Operations must confirm the supported route before staging.</p><textarea id="request-text" aria-label="Hosting request template">${helpRequest}\nAsset type, final files, captions, and placement:\nPublic-release approval / third-party permissions:</textarea><button class="button" data-copy-request>Copy request outline</button>`,
  profile:()=>`<h2 id="dialog-title">Set up your author profile</h2><ol><li>Sign in to Microsoft Tech Community with your work account.</li><li>Open your avatar → <strong>My Settings → Personal</strong>. Add the required photo; a short bio is optional.</li><li>Open <strong>Profile</strong> and copy its URL.</li><li>Send your <strong>username and profile URL to Katie Zoller</strong> in the article’s Teams chat. Request Blog Author and Blog Editor rights.</li><li>Confirm the main author before staging. That person creates and saves the platform draft shell.</li></ol><p>All authors need profiles. Profile permissions may take time, so complete this before your article is ready to stage.</p><p class="notice">Being able to edit the Word document does not automatically grant access to the Tech Community draft.</p><a class="button" href="${links.guide}">Read the detailed instructions</a>`,
  taxonomy:()=>`<h2 id="dialog-title">Tags & publishing metadata</h2><p>Suggest <strong>one to four tags</strong> from the Discovery taxonomy in your article’s Teams chat. Operations reviews them as part of SEO preparation.</p><dl><dt>Tags</dt><dd>Categories that help readers find related work. Choose from the approved taxonomy.</dd><dt>SEO description</dt><dd>Search-oriented description. Review it with Operations separately from the article body.</dd><dt>Thumbnail summary</dt><dd>A short introduction below the article thumbnail, no more than 140 characters.</dd><dt>Figure caption</dt><dd>Explains the figure and its significance, with credit where required.</dd><dt>Alt text</dt><dd>Describes meaningful visual content for readers using assistive technology.</dd></dl><p class="notice">The taxonomy workbook has not yet been identified among the supplied links.</p>${suppliedResources()}`,
  'article-template':()=>`<h2 id="dialog-title">Start the article in shared files</h2><p>Ask Operations for the approved Word template, if available, and save it in your article workspace. Keep drafts out of personal OneDrive so the team can retain access.</p><h3>Prepare before drafting</h3><ul><li>A descriptive title that includes “Microsoft Discovery.”</li><li>The intended scientific audience and the practical takeaway.</li><li>A repeatable workflow, supporting references, and publicly releasable evidence.</li><li>Approved figures, captions, alt text, and related-content links.</li><li>The main author, co-authors, and any third-party approvals needed.</li></ul><p class="notice">The approved article-template file has not been provided or identified yet. This checklist is preparation guidance, not a replacement template.</p>${externalLink(links.tracker,'Find your article workspace','button')}`,
  'asset-templates':()=>`<h2 id="dialog-title">Asset templates</h2><p>Confirm the program’s approved asset templates with Operations before starting a new design. Use the agreed source files for thumbnails, banners, and supporting visuals.</p><dl><dt>Thumbnail canvas</dt><dd>800 × 400 px. Export to PNG or JPG.</dd><dt>Header canvas</dt><dd>1300 × 500 px. Export to PNG or JPG.</dd><dt>Visual review</dt><dd>Coordinate with Operations and August Laguio for design and amplification needs.</dd></dl><p class="notice">The supplied presentation has not been verified as the approved asset template. Its owner and last updated date are also unconfirmed.</p>${externalLink(links.supplied[2].url,'Open supplied presentation','button')}<p style="margin-top:16px"><button class="text-link" data-open="asset-guide">Review all asset requirements</button></p>`,
};
const dialog=$('#resource-dialog');
let dialogTrigger=null;
function openResource(name,trigger){
  if(!resources[name])return;
  if(!dialog.open)dialogTrigger=trigger||document.activeElement;
  $('#dialog-body').innerHTML=resources[name]();
  if(!dialog.open)dialog.showModal();
  dialog.scrollTop=0;$('#close-dialog').focus();
}
function closeResource(){dialog.close();dialogTrigger?.focus();}
$('#close-dialog').addEventListener('click',closeResource);
dialog.addEventListener('cancel',event=>{event.preventDefault();closeResource();});
dialog.addEventListener('click',event=>{if(event.target===dialog){const b=dialog.getBoundingClientRect();if(event.clientX<b.left||event.clientX>b.right||event.clientY<b.top||event.clientY>b.bottom)closeResource();}});
document.addEventListener('click',async event=>{
  const opener=event.target.closest('[data-open]');
  if(opener){openResource(opener.dataset.open,opener);return;}
  if(event.target.closest('[data-copy-request]')){
    const field=$('#request-text');
    try{await navigator.clipboard.writeText(field.value);$('#toast').textContent='Request outline copied. Add your details in Teams.';$('#toast').hidden=false;setTimeout(()=>$('#toast').hidden=true,4000);}
    catch{field.focus();field.select();event.target.textContent='Text selected — copy to continue';}
  }
});
document.querySelectorAll('.page-nav a').forEach(a=>a.addEventListener('click',()=>{document.querySelectorAll('.page-nav a').forEach(link=>link.classList.toggle('active',link===a));}));
