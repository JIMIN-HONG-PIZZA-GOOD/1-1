from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
from playwright.sync_api import sync_playwright
ROOT=Path('/mnt/data/class-bulletin-archive')
poster=ROOT/'tests/test-poster.jpg'
im=Image.new('RGB',(1000,1350),'#e7e8ec')
d=ImageDraw.Draw(im)
d.polygon([(100,75),(905,155),(790,1250),(140,1190)],fill='#fffef7',outline='#82849e',width=12)
font='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
f=ImageFont.truetype(font,61)
d.text((195,240),'TEST SCHOOL POSTER',font=f,fill='#202548')
for i in range(8):d.line([(215,415+i*86),(700,430+i*83)],fill='#747b91',width=11)
im.save(poster,quality=88)
html=(ROOT/'preview/demo.html').read_text()
with sync_playwright() as p:
 for device,width,height in [('desktop',1280,880),('mobile',390,844)]:
  b=p.chromium.launch(executable_path='/usr/lib/chromium/chromium',headless=True,args=['--no-sandbox','--disable-dev-shm-usage','--disable-gpu','--no-zygote','--single-process','--disable-background-networking'])
  page=b.new_page(viewport={'width':width,'height':height},device_scale_factor=1)
  errors=[]
  page.on('pageerror',lambda err: errors.append(str(err)))
  page.set_content(html,wait_until='domcontentloaded')
  page.wait_for_timeout(350)
  assert page.locator('#homeView').is_visible(), f'home not visible {device}'
  assert page.locator('#homePosts .post-card').count()==3
  if device=='mobile':
   page.screenshot(path=str(ROOT/'preview/mobile.png'),full_page=True)
  # archive calendar filtering
  page.locator('[data-view="archive"]').first.click()
  assert page.locator('#archiveView').is_visible()
  page.locator('[data-mode="day"]').click()
  assert page.locator('#calendarPanel').is_visible()
  page.locator('[data-mode="all"]').click()
  assert page.locator('#archivePosts .post-card').count()==4
  page.locator('#searchInput').fill('글쓰기')
  assert page.locator('#archivePosts .post-card').count()==1
  page.locator('#clearFilters').click()
  assert page.locator('#archivePosts .post-card').count()==4
  # post view
  page.locator('#archivePosts .post-card').first.click()
  assert page.locator('#detailModal').is_visible()
  page.locator('[data-action="close-modal"]').last.click(force=True)
  # upload and image transformation
  page.locator('[data-view="upload"]').first.click()
  page.locator('#galleryFile').set_input_files(str(poster))
  page.wait_for_timeout(1400)
  assert page.locator('#stageTwo').is_visible(), f'photo editing not visible {device}:'+str(errors)
  page.locator('#autoEnhance').click()
  if device=='desktop':page.screenshot(path=str(ROOT/'preview/editor.png'),full_page=True)
  page.locator('#brightness').fill('20')
  page.locator('#cropModeButton').click()
  assert page.locator('#cropArea').is_visible()
  page.locator('#cropModeButton').click()
  page.locator('#toggleMask').click()
  mask=page.locator('#maskCanvas')
  r=mask.bounding_box()
  page.mouse.move(r['x']+r['width']*.35,r['y']+r['height']*.4)
  page.mouse.down()
  page.mouse.move(r['x']+r['width']*.7,r['y']+r['height']*.49,steps=5)
  page.mouse.up()
  page.wait_for_timeout(850)
  page.locator('#toggleMask').click()
  page.locator('#toInfo').click()
  page.wait_for_timeout(900)
  assert page.locator('#stageThree').is_visible(), f'form not visible {device}: {errors}'
  assert page.locator('#finalPreview').get_attribute('src').startswith('data:image/jpeg;base64,')
  page.locator('#submitPostForm input[name=title]').fill('학교 게시물 테스트')
  page.locator('#submitPostForm input[name=privacyConfirmed]').check()
  page.locator('#savePostButton').click()
  page.wait_for_timeout(3500)
  print('after submit',device,'toast:',page.locator('#toast').inner_text(),'stage:',page.locator('#stageThree').is_visible())
  assert page.locator('#homeView').is_visible(), f'no home after save {device}: {errors}'
  assert page.locator('#homePosts .post-card').count()==3
  page.locator('[data-view="admin"]').first.click()
  page.wait_for_timeout(300)
  assert page.locator('#usersList .user-row').count()==2
  print(device,'OK','browser errors',errors)
  page.close()
  b.close()
